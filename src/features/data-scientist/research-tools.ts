/**
 * The Data Scientist's agent: it answers business questions from the research's semantic models in an
 * analysis text and a small workbook — pivot sheets (rows, columns, values, filters; the DAX is written
 * by code), charts of a pivot's figures, and calculation sheets (Excel formulas reading the pivots with
 * GETPIVOTDATA), so a figure the models do not have is computed in the open, as the user would in Excel.
 */
import { z } from "zod";
import { datasetText, withContexts } from "@/agent/context";
import { tool, type Agent } from "@/agent/loop";
import { daxTool, filterOp, missingValues, queryTool } from "@/agent/pbi-tools";
import { modelTools } from "@/agent/model-tools";
import { rowsText } from "@/services/pbi";
import { CHARTS, pivot, readResearch, resultName, runSheet, sheetDax, sheetName, type CalcSheet, type ChartSheet, type PivotSheet, type ResearchEntry, type Sheet } from "@/features/data-scientist/research";
import { columnName, computeWorkbook } from "@/features/data-scientist/formulas";
import { patch, type DatasetEntry } from "@/lib/store";
import { researchInstructions } from "@/prompts/data-scientist";

const sheetId = () => crypto.randomUUID().slice(0, 8);
/** Fields: a list (a single field, or several separated by commas, are accepted too). */
const fieldList = z.preprocess((v) => (typeof v === "string" ? (v.trim() ? v.split(/,(?![^[]*\])/).map((x) => x.trim()) : []) : v), z.array(z.string()).default([]));
const filterList = z.array(z.object({ field: z.string().describe("A column 'T'[col]"), op: filterOp.optional(), values: z.preprocess((v) => (v === undefined || Array.isArray(v) ? v : [v]), z.array(z.union([z.string(), z.number(), z.boolean()]))).optional(), value: z.union([z.string(), z.number()]).optional(), to: z.union([z.string(), z.number()]).optional() })).default([]);

/** A sheet as the AI sees it in the state: its kind, parameters and size (pivots: the names GETPIVOTDATA takes). */
function sheetText(s: Sheet, name: (id: string) => string) {
  if (s.kind === "calc") return `- "${s.name}" calc — ${s.cells.length} rows × ${Math.max(0, ...s.cells.map((r) => r.length))} columns`;
  const spec = s.spec;
  const fields = spec ? `rows ${spec.rows.join(", ") || "—"}; columns ${spec.columns.join(", ") || "—"}; values ${spec.values.join(", ")}${spec.filters.length ? `; ${spec.filters.length} filter(s)` : ""}` : s.kind === "chart" ? `DAX x ${s.x}, y ${s.y?.join(", ")}` : "";
  const names = spec ? ` — GETPIVOTDATA values: ${spec.values.map((v) => `"${resultName(v)}"`).join(", ")}; fields: ${[...spec.rows, ...spec.columns].map((f) => `"${resultName(f)}"`).join(", ") || "none"}` : "";
  return `- "${s.name}" ${s.kind === "chart" ? `${s.chart} chart` : "pivot"} on "${name(s.datasetId)}" — ${fields}${s.error ? ` — ERROR: ${s.error.slice(0, 120)}` : s.rows ? ` — ${s.rows.length} rows` : ""}${names}`;
}

/** The research as the AI sees it: models, the analysis text and the sheets. */
function researchText(r: ResearchEntry, datasets: DatasetEntry[]) {
  const name = (id: string) => datasets.find((d) => d.id === id)?.name ?? id;
  return [
    `Models: ${r.datasetIds.map((id) => `"${name(id)}" (dataset_id ${id})`).join(", ") || "none (the user chooses them in the Datasets tab)"}.`,
    `Analysis text (${r.text.length} characters):\n${r.text ? r.text.slice(0, 1500) : "empty"}`,
    `Sheets of the workbook (${r.sheets.length}, in order):\n${r.sheets.map((s) => sheetText(s, name)).join("\n") || "none"}`,
  ].join("\n\n");
}

/** A calculation sheet's results as text: each row, "A1=…" for formulas with their value, errors with why. */
function calcText(sheet: CalcSheet, sheets: Sheet[]) {
  const results = computeWorkbook(sheets).get(sheet.name) ?? [];
  const errors: string[] = [];
  const lines = results.map((line, r) =>
    line
      .map((c, k) => {
        const ref = `${columnName(k)}${r + 1}`;
        if (c.error) errors.push(`${ref}: ${c.error} ${c.why ?? ""}`);
        return sheet.cells[r][k]?.startsWith("=") ? `${ref}=${typeof c.value === "number" ? Math.round(c.value * 10_000) / 10_000 : c.value}` : String(c.value ?? "");
      })
      .join(" | "),
  );
  return { text: lines.slice(0, 30).join("\n"), errors };
}

/** The pivots a calc sheet reads only partly: "Sales by zone" has 11 rows, the sheet reads 1. */
function coverage(calc: CalcSheet, sheets: Sheet[]) {
  return sheets.flatMap((p) => {
    if (p.kind === "calc" || !p.spec || !p.rows?.length) return [];
    const quoted = `"${p.name.toLowerCase()}"`;
    const read = calc.cells.filter((line) => line.some((c) => /GETPIVOTDATA/i.test(c) && c.toLowerCase().includes(quoted))).length;
    const items = pivot(p.spec, p.rows).lines.length;
    return read && read < items ? [`"${p.name}" has ${items} rows, this sheet reads ${read}`] : [];
  });
}

export function researchAgent(entry: () => ResearchEntry, datasets: () => DatasetEntry[]): Agent {
  // The research as last written here: the screen's copy only catches up after a render, so tools run
  // in one reply (several sheets) read this one and never write over each other.
  let written: ResearchEntry | undefined;
  const get = () => {
    const shown = readResearch(entry());
    return written && written.id === shown.id && written.updatedAt >= shown.updatedAt ? written : shown;
  };
  const save = (change: Partial<ResearchEntry>) => {
    written = { ...get(), ...change, updatedAt: Math.max(Date.now(), get().updatedAt + 1) };
    return patch("researches", written.id, { ...change, updatedAt: written.updatedAt });
  };
  const models = () => datasets().filter((d) => get().datasetIds.includes(d.id));
  const model = (id?: string) => {
    const list = models();
    const d = (id ? list.find((x) => x.id === id || x.name.toLowerCase() === id.toLowerCase()) : undefined) ?? (list.length === 1 ? list[0] : undefined);
    if (!d) throw new Error(list.length ? `Give dataset_id: one of ${list.map((x) => `${x.id} (${x.name})`).join(", ")}.` : "No model in this research: ask the user to choose one in the Datasets tab.");
    return d;
  };
  /** The model the model tools read (their dataset_id, or the only model). */
  let described: string | undefined;
  /** Puts a sheet in the workbook (replacing the one of the same name, in its place; else after `after`, else at the end). */
  const place = async (s: Sheet, after?: string) => {
    const sheets = get().sheets.filter((x) => x.name.toLowerCase() !== s.name.toLowerCase());
    const old = get().sheets.findIndex((x) => x.name.toLowerCase() === s.name.toLowerCase());
    const anchor = after ? sheets.findIndex((x) => x.name.toLowerCase() === after.toLowerCase()) : -1;
    sheets.splice(old >= 0 ? old : anchor >= 0 ? anchor + 1 : sheets.length, 0, s);
    await save({ sheets });
  };
  /** Runs a data sheet's query, saves it and says what it returned (and why nothing, when a filter value does not exist). */
  const putData = async (s: PivotSheet | ChartSheet, after?: string) => {
    const d = model(s.datasetId);
    const done = await runSheet(s, { id: d.id, groupId: d.groupId, info: d.info });
    if (done.error) throw new Error(`${done.error}\nDAX:\n${sheetDax(s, d.info)}`);
    await place(done, after);
    const empty = !done.rows?.length && s.spec ? await missingValues({ id: d.id, groupId: d.groupId }, { values: s.spec.values, group_by: [...s.spec.rows, ...s.spec.columns], filters: s.spec.filters }, d.info) : "";
    const p = s.spec && done.rows?.length ? pivot(s.spec, done.rows) : undefined;
    return `Sheet "${s.name}" (${s.kind}): ${rowsText(done.rows ?? [], 12)}${empty}${p ? `\nGETPIVOTDATA on it: values ${s.spec!.values.map((v) => `"${resultName(v)}"`).join(", ")}; fields ${[...s.spec!.rows, ...s.spec!.columns].map((f) => `"${resultName(f)}"`).join(", ") || "none"}.` : ""}`;
  };
  const pivotArgs = {
    name: z.string().describe("Sheet name (tab), short: it is how calc formulas reach it"),
    dataset_id: z.string().optional(),
    rows: fieldList.describe("Columns grouped as rows, e.g. ['Product'[Category]]"),
    columns: fieldList.describe("Columns grouped as columns (e.g. 'Date'[Year]), often empty"),
    values: fieldList.describe("Measures [M], or Sum('T'[col]) / Avg / Count / DistinctCount / Min / Max"),
    filters: filterList.describe("Filters on columns (check the values exist first)"),
    after: z.string().optional().describe("Sheet to put it after (default: at the end)"),
  };
  return {
    // Every request is explored, then confirmed by the user, then shown in the workbook.
    answerOnScreen: true,
    instructions: researchInstructions(),
    context: async (focus) => withContexts(await Promise.all(models().map(async (d) => `${await datasetText(d, focus)}\n(dataset_id ${d.id})`)), models().map((d) => d.contextIds)),
    state: async () => researchText(get(), datasets()),
    // A pivot backs every answer (the figures behind it) — the one answering the request, never one for show.
    beforeAnswer: () =>
      models().length && !get().sheets.some((s) => s.kind === "pivot" && !s.error && s.rows?.length)
        ? "Add the pivot sheet (add_pivot) whose figures answer the request before answering — never an unrelated one. If no pivot of the models can answer it, call ask_user with what you found and the real options instead."
        : undefined,
    tools: [
      daxTool(() => undefined, () => models()),
      queryTool(() => undefined, () => models()),
      // The model's tables, columns and measures in detail (its hidden tables too: what can rebuild a figure).
      ...modelTools(() => model(described)).map((t) => ({
        ...t,
        args: t.args.extend({ dataset_id: z.string().optional().describe("Model id (when there are several)") }),
        run: (args: Record<string, unknown>) => ((described = args.dataset_id as string | undefined), t.run(args)),
      })),
      tool({
        name: "write_text",
        description: "Writes the analysis text shown above the workbook (markdown, the whole text: it replaces the current one).",
        args: z.object({ text: z.string() }),
        run: async ({ text }) => (await save({ text }), "Analysis text written."),
      }),
      tool({
        name: "add_pivot",
        description: "Adds a pivot sheet (or replaces the sheet of that name): its query runs at once; its first rows and the names GETPIVOTDATA takes are returned.",
        args: z.object(pivotArgs),
        run: async ({ name, dataset_id, rows, columns, values, filters, after }) => {
          if (!values.length) throw new Error('Give "values" (measures).');
          const d = model(dataset_id);
          return putData({ id: sheetId(), kind: "pivot", name: sheetName(name), datasetId: d.id, spec: { rows, columns, values, filters } }, after);
        },
      }),
      tool({
        name: "add_chart",
        description: "Adds a chart sheet (or replaces the sheet of that name) with a pivot's parameters: rows = the categories (x, or the dates), columns = the series (often empty), values = what is measured.",
        args: z.object({ ...pivotArgs, chart: z.enum(CHARTS).describe("column, bar (ranking), line / area (trend over time), pie (few parts of a whole), scatter (two values)") }),
        run: async ({ name, dataset_id, rows, columns, values, filters, after, chart }) => {
          if (!values.length) throw new Error('Give "values" (measures).');
          const d = model(dataset_id);
          return putData({ id: sheetId(), kind: "chart", chart, name: sheetName(name), datasetId: d.id, spec: { rows, columns, values, filters } }, after);
        },
      }),
      tool({
        name: "add_calc",
        description:
          'Adds a calculation sheet (or replaces the sheet of that name): rows of cells, each a text, a number or an Excel formula ("=…"). Read a pivot sheet with GETPIVOTDATA("value", "pivot sheet name", "field", "item", …) — never by cell; other cells by A1, B2:B5 or \'Calc name\'!B2. Functions: SUM AVERAGE MIN MAX COUNT ROUND ABS IF IFERROR AND OR NOT. The computed values (and any error) are returned.',
        args: z.object({
          name: z.string().describe("Sheet name (tab)"),
          cells: z.array(z.array(z.preprocess((v) => (v === null || v === undefined ? "" : String(v)), z.string()))).min(1).describe('Rows of cells, e.g. [["Region", "Sales", "Target", "Gap %"], ["North", "=GETPIVOTDATA(\\"Sales\\", \\"Sales 2025\\", \\"Region\\", A2)", "=GETPIVOTDATA(\\"Target\\", \\"Targets\\", \\"Region\\", A2)", "=B2/C2-1"]]'),
          note: z.string().optional().describe("How it is computed, in a sentence"),
          after: z.string().optional(),
        }),
        run: async ({ name, cells, note, after }) => {
          const s: CalcSheet = { id: sheetId(), kind: "calc", name: sheetName(name), cells, note };
          await place(s, after);
          const { text, errors } = calcText(s, get().sheets);
          const partial = coverage(s, get().sheets);
          return `Sheet "${s.name}" (calc) computed:\n${text}${errors.length ? `\nErrors to fix (add_calc again with the same name): ${errors.slice(0, 6).join("; ")}` : ""}${partial.length ? `\nOnly part of the pivots is read: ${partial.join("; ")}. Add a row per item (add_calc again with the same name) unless the user asked for these only.` : ""}`;
        },
      }),
      tool({
        name: "delete_sheet",
        description: "Removes a sheet from the workbook.",
        args: z.object({ name: z.string() }),
        run: async ({ name }) => {
          if (!get().sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) throw new Error(`No sheet "${name}". Sheets: ${get().sheets.map((s) => s.name).join(", ") || "none"}.`);
          await save({ sheets: get().sheets.filter((s) => s.name.toLowerCase() !== name.toLowerCase()) });
          return `Sheet "${name}" removed.`;
        },
      }),
    ],
  };
}
