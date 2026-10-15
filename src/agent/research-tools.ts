/**
 * Researcher: the AI builds analysis sheets in an Excel workbook from several semantic models —
 * PivotTables connected to a model (create_table_from_tcd) and tables whose columns are formulas — and
 * charts of them. It never types data: the values are computed by Excel (and the models). Changes are
 * kept in the research and published when the answer is done (or when values are read).
 */
import { z } from "zod";
import { dax, rowsText } from "@/lib/pbi";
import { resolveField, modelSummary } from "@/lib/model";
import { sheetValues } from "@/lib/research";
import type { DatasetEntry, ResearchEntry } from "@/lib/store";
import { CHART_TYPES, sheetName, tableName, type FormulaSheet, type PivotSheet, type ResearchSheet } from "@/lib/workbook";
import { tool, type Tool } from "./loop";
import { runSandboxed } from "./sandbox";

type Ctx = {
  research: () => ResearchEntry | undefined;
  datasets: () => DatasetEntry[];
  /** Saves the sheets (in the library), marked to publish. */
  setSheets: (sheets: ResearchSheet[]) => Promise<void>;
  /** Publishes now when there are unpublished changes (to read computed values). */
  publish: () => Promise<void>;
};

/** The lists the AI gets with every message: the models it may use and the workbook's sheets. */
export function researchState(r: ResearchEntry | undefined, datasets: DatasetEntry[], dirty: boolean) {
  if (!r) return "No research open.";
  const sheet = (s: ResearchSheet) => {
    if (s.kind === "pivot") return `- PIVOT "${s.name}" on model ${s.dataset.name}: rows ${s.rows.map((f) => `'${f.table}'[${f.column}]`).join(", ") || "none"}; columns ${s.columns.map((f) => `'${f.table}'[${f.column}]`).join(", ") || "none"}; values ${s.values.map((v) => `[${v.measure}]`).join(", ")}${s.filters.length ? `; filters ${s.filters.map((f) => `'${f.table}'[${f.column}]${f.value !== undefined ? ` = ${f.value}` : ""}`).join(", ")}` : ""}`;
    if (s.kind === "table") return `- TABLE "${s.name}" (Excel table ${tableName(s.name)}, ${s.rows} rows): ${s.columns.map((c) => `${c.name}${c.formula ? ` = ${c.formula}` : " (empty)"}`).join(" | ")}`;
    return `- CHART "${s.name}" (${s.type}) of "${s.source}"`;
  };
  return [
    `Models (list_datasets): ${datasets.map((d) => `${d.name} (dataset_id ${d.id})`).join("; ") || "none: the user must add one to the research"}`,
    `Sheets of the workbook (list_tables), one analysis per sheet:\n${r.sheets.map(sheet).join("\n") || "none yet"}`,
    dirty ? "Some changes are not published yet (published at the end of your answer, or when you read values)." : "",
  ]
    .filter(Boolean)
    .join("\n\n");
}

/** What the AI knows of the models (once, then when they change). */
export const datasetsContext = (datasets: DatasetEntry[]) => datasets.map((d) => `## Model "${d.name}" (dataset_id ${d.id})\n${d.info ? modelSummary(d.info) : d.error ?? "Structure not read."}`).join("\n\n");

const fieldList = z.array(z.string()).describe("Columns of the model: 'Table'[Column]");

export function researchTools(ctx: Ctx): Tool[] {
  const research = () => {
    const r = ctx.research();
    if (!r) throw new Error("Open a research first.");
    return r;
  };
  const datasetOf = (id: string) => {
    const d = ctx.datasets().find((x) => x.id === id || x.name.toLowerCase() === id.toLowerCase());
    if (!d) throw new Error(`No model "${id}" in this research. Models: ${ctx.datasets().map((x) => `${x.name} (${x.id})`).join(", ") || "none"}.`);
    return d;
  };
  const sheetOf = (name: string) => {
    const s = research().sheets.find((x) => x.name.toLowerCase() === name.toLowerCase() || tableName(x.name).toLowerCase() === name.toLowerCase());
    if (!s) throw new Error(`No sheet "${name}". Sheets: ${research().sheets.map((x) => x.name).join(", ") || "none"}.`);
    return s;
  };
  const tableOf = (name: string) => {
    const s = sheetOf(name);
    if (s.kind !== "table") throw new Error(`"${s.name}" is a ${s.kind}, not a formula table.`);
    return s;
  };
  const free = (name: string) => {
    const n = sheetName(name);
    if (research().sheets.some((s) => s.name.toLowerCase() === n.toLowerCase())) throw new Error(`A sheet "${n}" exists already.`);
    return n;
  };
  const replace = (old: ResearchSheet, next: ResearchSheet) => ctx.setSheets(research().sheets.map((s) => (s === old ? next : s)));
  /** A pivot's fields checked against its model: columns for rows / columns / filters, measures for values. */
  const pivot = (name: string, a: { dataset_id: string; rows?: string[]; columns?: string[]; values: string[]; filters?: { field: string; value?: string }[] }): PivotSheet => {
    const d = datasetOf(a.dataset_id);
    const column = (spec: string) => {
      const f = resolveField(d.info, spec);
      if (f.kind === "measure") throw new Error(`[${f.name}] is a measure: put it in values.`);
      return { table: f.table, column: f.name };
    };
    const values = a.values.map((spec) => {
      const f = resolveField(d.info, spec);
      if (f.kind !== "measure") throw new Error(`'${f.table}'[${f.name}] is a column: values must be measures of the model (an aggregation of a column needs a measure: ask for one in the Dataset editor).`);
      return { table: f.table, measure: f.name };
    });
    if (!values.length) throw new Error("A pivot needs at least one measure in values.");
    return { kind: "pivot", name, dataset: { id: d.id, name: d.name }, rows: (a.rows ?? []).map(column), columns: (a.columns ?? []).map(column), values, filters: (a.filters ?? []).map((f) => ({ ...column(f.field), value: f.value })) };
  };
  const pivotArgs = {
    dataset_id: z.string(),
    rows: fieldList.optional(),
    columns: fieldList.optional(),
    values: z.array(z.string()).describe("Measures of the model: [Measure]"),
    filters: z.array(z.object({ field: z.string(), value: z.string().optional().describe("One value kept (as in the data); none: the filter is shown, all selected") })).optional(),
  };

  return [
    tool({
      name: "create_table_from_tcd",
      description: "Adds a sheet with a PivotTable connected to a model (the real Excel PivotTable, computed by Excel from the model): rows, columns, filters = columns of the model; values = measures.",
      args: z.object({ name: z.string(), ...pivotArgs }),
      run: async ({ name, ...a }) => {
        const n = free(name);
        await ctx.setSheets([...research().sheets, pivot(n, a)]);
        return `Pivot "${n}" added (computed when the workbook is published).`;
      },
    }),
    tool({
      name: "update_table_from_tcd",
      description: "Replaces the fields of a pivot (all of them).",
      args: z.object({ name: z.string(), ...pivotArgs }),
      run: async ({ name, ...a }) => {
        const s = sheetOf(name);
        if (s.kind !== "pivot") throw new Error(`"${s.name}" is not a pivot.`);
        await replace(s, pivot(s.name, a));
        return `Pivot "${s.name}" changed.`;
      },
    }),
    tool({
      name: "create_table",
      description: "Adds a sheet with an empty Excel table (its columns are then formulas: create_column). rows = number of rows (each column's formula fills them).",
      args: z.object({ name: z.string(), rows: z.number().int().min(1).max(5000) }),
      run: async ({ name, rows }) => {
        const n = free(name);
        await ctx.setSheets([...research().sheets, { kind: "table", name: n, rows, columns: [] }]);
        return `Table "${n}" added (Excel table name: ${tableName(n)}).`;
      },
    }),
    tool({
      name: "create_column",
      description: 'Adds a column to a formula table. formula = an Excel formula for one row, e.g. =ROW()-1, =[@Price]*[@Quantity], =XLOOKUP([@Code], Other[Code], Other[Name]), =INDEX(Other[Name], [@N]). Never type values.',
      args: z.object({ table: z.string(), name: z.string(), formula: z.string().optional() }),
      run: async ({ table, name, formula }) => {
        const t = tableOf(table);
        if (t.columns.some((c) => c.name.toLowerCase() === name.toLowerCase())) throw new Error(`"${t.name}" has a column "${name}" already.`);
        await replace(t, { ...t, columns: [...t.columns, { name, formula: formula && `=${formula.replace(/^=/, "")}` }] });
        return `Column "${name}" added to "${t.name}".`;
      },
    }),
    tool({
      name: "update_column",
      description: "Changes the formula of a column of a formula table.",
      args: z.object({ table: z.string(), name: z.string(), formula: z.string() }),
      run: async ({ table, name, formula }) => {
        const t = tableOf(table);
        if (!t.columns.some((c) => c.name.toLowerCase() === name.toLowerCase())) throw new Error(`No column "${name}" in "${t.name}".`);
        await replace(t, { ...t, columns: t.columns.map((c) => (c.name.toLowerCase() === name.toLowerCase() ? { ...c, formula: `=${formula.replace(/^=/, "")}` } : c)) } as FormulaSheet);
        return `Formula of "${name}" changed.`;
      },
    }),
    tool({
      name: "rename_column",
      description: "Renames a column of a formula table (formulas using the old name must be changed too).",
      args: z.object({ table: z.string(), name: z.string(), new_name: z.string() }),
      run: async ({ table, name, new_name }) => {
        const t = tableOf(table);
        await replace(t, { ...t, columns: t.columns.map((c) => (c.name.toLowerCase() === name.toLowerCase() ? { ...c, name: new_name } : c)) });
        return `Column renamed to "${new_name}".`;
      },
    }),
    tool({
      name: "delete_column",
      description: "Removes a column of a formula table.",
      args: z.object({ table: z.string(), name: z.string() }),
      run: async ({ table, name }) => {
        const t = tableOf(table);
        await replace(t, { ...t, columns: t.columns.filter((c) => c.name.toLowerCase() !== name.toLowerCase()) });
        return `Column "${name}" removed.`;
      },
    }),
    tool({
      name: "create_chart",
      description: `Adds a sheet with a chart of another sheet (a pivot or a formula table). Types: ${CHART_TYPES.join(", ")}.`,
      args: z.object({ name: z.string(), source: z.string().describe("Name of the sheet charted"), type: z.enum(CHART_TYPES) }),
      run: async ({ name, source, type }) => {
        const src = sheetOf(source);
        if (src.kind === "chart") throw new Error("A chart charts a table or a pivot, not another chart.");
        const n = free(name);
        await ctx.setSheets([...research().sheets, { kind: "chart", name: n, source: src.name, type }]);
        return `Chart "${n}" added.`;
      },
    }),
    tool({
      name: "delete_table",
      description: "Deletes a sheet (pivot, table or chart).",
      args: z.object({ name: z.string() }),
      run: async ({ name }) => {
        const s = sheetOf(name);
        await ctx.setSheets(research().sheets.filter((x) => x !== s && !(x.kind === "chart" && x.source === s.name)));
        return `Sheet "${s.name}" deleted.`;
      },
    }),
    tool({
      name: "get_table_values",
      description: "The values Excel computed for a sheet (first rows): to check a pivot or formulas. Publishes the workbook first when needed.",
      args: z.object({ table: z.string(), limit: z.number().optional() }),
      readOnly: true,
      run: async ({ table, limit = 30 }) => {
        const s = sheetOf(table);
        await ctx.publish();
        const { values, total, address } = await sheetValues(research(), s.name, limit);
        if (!values.length) return `"${s.name}" is empty.`;
        const [head, ...rows] = values;
        return `${address} (${total - 1} rows)\n${rowsText(rows.map((r) => Object.fromEntries((head as unknown[]).map((h, i) => [String(h || `col${i + 1}`), r[i]]))), limit)}`;
      },
    }),
    tool({
      name: "run_dax",
      description: "Runs a DAX query (EVALUATE …) on one of the models (a thinking tool: to look at values, test a measure; max 50 rows shown).",
      args: z.object({ dataset_id: z.string(), query: z.string() }),
      readOnly: true,
      run: async ({ dataset_id, query }) => {
        const d = datasetOf(dataset_id);
        return rowsText(await dax({ id: d.id, groupId: d.groupId }, query));
      },
    }),
    tool({
      name: "run_js",
      description: 'Runs JavaScript on values of sheets (no network, 5 s): `data` = {sheet name: rows as arrays, first row = header}. Return a value (or console.log). To compute, never to write.',
      args: z.object({ code: z.string().describe("Body of an async function; e.g. return data.Sales.slice(1).reduce((s, r) => s + r[1], 0)"), sheets: z.array(z.string()).describe("Sheets whose values are given") }),
      readOnly: true,
      run: async ({ code, sheets }) => {
        await ctx.publish();
        const data: Record<string, unknown[][]> = {};
        for (const name of sheets) data[sheetOf(name).name] = (await sheetValues(research(), sheetOf(name).name, 5000)).values;
        const out = await runSandboxed(code, data);
        if (out.error) throw new Error(`${out.error}${out.log.length ? `\nlog: ${out.log.join("\n")}` : ""}`);
        return [out.log.length && `log:\n${out.log.join("\n").slice(0, 2000)}`, `result: ${out.result ?? "undefined"}`].filter(Boolean).join("\n").slice(0, 6000);
      },
    }),
  ];
}
