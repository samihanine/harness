/**
 * Dataset editor: the AI changes a semantic model's structure — columns, measures, relationships,
 * Power Query steps — in a draft of its definition (TMDL) saved when its answer is done. DAX is
 * evaluated on the model before it is written. The data itself changes with a refresh, started by the
 * user from the screen (never by the AI: about 8 refreshes a day on Pro).
 */
import { z } from "zod";
import { dax, rowsText, type Ref } from "@/services/pbi";
import { fromTmdl, resolveField, type ModelInfo } from "@/services/model";
import { readModelFiles, writeModelFiles } from "@/services/tmdl";
import { ModelDraft, STEP_KINDS, daxRef, dynamicSource, imageSources, parseTable, type ModelFile, type Step } from "@/features/datasets/tmdl-edit";
import { graph } from "@/services/ms";
import { tool, type Tool } from "@/agent/loop";

type State = { draft: ModelDraft; base: ModelFile[]; changes: string[] };
const drafts = new Map<string, State>();
const previous = new Map<string, ModelFile[]>();

/** The model's draft, made on first use from its definition. */
async function open(ref: Ref) {
  let state = drafts.get(ref.id);
  if (!state) {
    const base = await readModelFiles(ref.id, ref.groupId);
    drafts.set(ref.id, (state = { draft: new ModelDraft(base), base, changes: [] }));
  }
  return state;
}

export const modelDraft = (ref: Ref) => drafts.get(ref.id);
/** The model with the draft's unsaved changes (e.g. a field parameter just added), or undefined without a draft. */
export const draftInfo = (ref: Ref): ModelInfo | undefined => {
  const state = drafts.get(ref.id);
  return state?.changes.length ? fromTmdl(state.draft.toFiles()) : undefined;
};

/**
 * create_field_parameter: a table of fields a slicer switches in visuals (measures or columns), in the
 * model's draft (saved with the screen's other changes; its values come with the next refresh).
 */
export const fieldParameterTool = (dataset: () => Ref | undefined): Tool =>
  tool({
    name: "create_field_parameter",
    description: "Creates a field parameter in the semantic model: a slicer on it lets the user switch the measure (or the axis column) of the visuals using it. Then use 'Name'[Name] as a slicer's field and in a visual's role instead of the field. Its values come after the user refreshes the data.",
    args: z.object({ name: z.string().describe('Table name, e.g. "Metric"'), fields: z.array(z.string()).min(2).describe("Measures [M] or columns 'T'[col], in order"), labels: z.array(z.string()).optional().describe("Names shown in the slicer (default: the field names)") }),
    run: async ({ name, fields, labels }) => {
      const r = dataset();
      if (!r) throw new Error("No semantic model.");
      const state = await open(r);
      const info = fromTmdl(state.draft.toFiles());
      const items = fields.map((f, i) => {
        const known = resolveField(info, f);
        return { label: labels?.[i] ?? known.name, ref: daxRef(known.table, known.name) };
      });
      const text = state.draft.addFieldParameter(name, items);
      state.changes.push(`field parameter '${name}' created`);
      return `${text} (in the model's draft; saved at the end of your answer)`;
    },
  });
export const discardModelDraft = (ref: Ref) => void drafts.delete(ref.id);

/**
 * Saves the draft when it has changes: refused when the model changed meanwhile (read again, the
 * changes are lost: the AI redoes them). Returns a summary, or undefined when nothing changed.
 */
export async function saveModelDraft(ref: Ref) {
  const state = drafts.get(ref.id);
  if (!state?.changes.length) return undefined;
  const fresh = await readModelFiles(ref.id, ref.groupId);
  const same = (a: ModelFile[], b: ModelFile[]) => a.length === b.length && a.every((f) => b.find((x) => x.path === f.path)?.text === f.text);
  if (!same(fresh, state.base)) {
    drafts.delete(ref.id);
    throw new Error("The model was changed by someone else meanwhile: your changes were not saved. Ask again.");
  }
  await writeModelFiles(ref.id, state.draft.toFiles(), ref.groupId);
  previous.set(ref.id, state.base);
  drafts.delete(ref.id);
  return `Saved to the model: ${state.changes.length} change(s). Refresh the data (button "Refresh the data") for new columns and Power Query steps to get their values.`;
}

/** The structure the AI gets with every message (list_tables): tables, columns, measures, steps, relationships. */
export async function modelState(ref: Ref | undefined, refresh?: string) {
  if (!ref) return "No dataset open.";
  const { draft, changes } = await open(ref);
  const info = fromTmdl(draft.toFiles());
  const lines = info.tables.map((t) => {
    let steps = "";
    try {
      steps = draft.query(t.name).steps?.list.map((s) => s.name).join(" → ") ?? "";
    } catch {}
    return [
      `- '${t.name}'${t.kind ? ` (${t.kind})` : ""}${t.hidden ? " hidden" : ""}`,
      `  columns: ${t.columns.map((c) => `${c.name} ${c.type ?? "?"}${c.expression ? " (DAX)" : ""}${c.hidden ? " hidden" : ""}${c.format ? ` "${c.format}"` : ""}`).join("; ") || "none"}`,
      t.measures.length ? `  measures: ${t.measures.map((m) => `[${m.name}] = ${String(m.expression ?? "").replace(/\s+/g, " ").slice(0, 100)}`).join("; ")}` : "",
      steps ? `  Power Query steps: ${steps}` : "",
    ]
      .filter(Boolean)
      .join("\n");
  });
  const rels = draft.listRelationships();
  const dynamic = info.tables.filter((t) => {
    try {
      return dynamicSource(draft.query(t.name).m);
    } catch {
      return false;
    }
  });
  return [
    `Tables (list_tables)${changes.length ? " — with your unsaved changes" : ""}:\n${lines.join("\n")}`,
    `Relationships: ${rels.map((r) => `${r.from} (${r.fromCardinality}) → ${r.to} (${r.toCardinality})${r.both ? " both directions" : ""}${r.active ? "" : " inactive"}`).join("; ") || "none"}`,
    dynamic.length ? `⚠ Power BI cannot refresh ${dynamic.map((t) => `'${t.name}'`).join(", ")}: a step reads a source whose address comes from a column (DynamicDataSourcesIsNotSupportedForRefresh). Remove that step (remove_last_step when it is the last) and use step_image_base64 for images.` : "",
    refresh ? `Data refresh (started by the user): ${refresh}` : "",
    changes.length ? `Unsaved changes (saved at the end of your answer): ${changes.join("; ")}` : "",
  ]
    .filter(Boolean)
    .join("\n\n");
}

const DATA_TYPES = ["string", "int64", "double", "decimal", "dateTime", "boolean"] as const;
const mType = z.enum(["text", "number", "integer", "date", "datetime", "logical"]);
const value = z.union([z.string(), z.number()]);

export function modelEditorTools(dataset: () => Ref | undefined): Tool[] {
  const ref = () => {
    const r = dataset();
    if (!r) throw new Error("Open a dataset first.");
    return r;
  };
  /** Runs a change on the draft and logs it (shown to the AI as pending). */
  const change = async (label: string, run: (d: ModelDraft) => string) => {
    const state = await open(ref());
    const text = run(state.draft);
    state.changes.push(label);
    return `${text} (in the draft; saved at the end of your answer)`;
  };
  /** DAX evaluated on the saved model before it is written (`check` false: references objects of the draft not saved yet). */
  const test = async (query: string, check = true) => {
    if (!check) return;
    await dax(ref(), query).catch((e) => {
      throw new Error(`${e instanceof Error ? e.message : e} — nothing written. (If it uses a column or measure you added in this answer, retry with check: false.)`);
    });
  };
  const step = (kind: Step["kind"], args: z.ZodRawShape, description: string) =>
    tool({
      name: `step_${kind}`,
      description: `Power Query step at the end of a table's query: ${description} The model's columns follow the step. The data changes after a refresh.`,
      args: z.object({ table: z.string(), name: z.string().optional().describe("Step name (optional)"), ...args }),
      run: async ({ table, ...rest }: any) => change(`step ${kind} on '${table}'`, (d) => d.addStep(table, { kind, ...rest } as Step)),
    });
  return [
    fieldParameterTool(dataset),
    tool({
      name: "create_measure",
      description: "Creates a measure (checked on the model first).",
      args: z.object({ table: z.string(), name: z.string(), expression: z.string(), format: z.string().optional().describe('e.g. "#,0", "0.0%", "Short Date"'), folder: z.string().optional(), description: z.string().optional(), check: z.boolean().optional() }),
      run: async ({ check, ...m }) => {
        await test(`DEFINE MEASURE ${daxRef(m.table, m.name)} = ${m.expression}\nEVALUATE ROW("value", ${daxRef(m.table, m.name)})`, check);
        return change(`measure [${m.name}] created`, (d) => d.setMeasure(m.table, m, true));
      },
    }),
    tool({
      name: "update_measure",
      description: "Changes a measure: expression (checked first), format, folder, description, hidden (only the given settings).",
      args: z.object({ table: z.string(), name: z.string(), expression: z.string().optional(), format: z.string().optional(), folder: z.string().optional(), description: z.string().optional(), hidden: z.boolean().optional(), check: z.boolean().optional() }),
      run: async ({ check, ...m }) => {
        if (m.expression) await test(`DEFINE MEASURE ${daxRef(m.table, m.name)} = ${m.expression}\nEVALUATE ROW("value", ${daxRef(m.table, m.name)})`, check);
        return change(`measure [${m.name}] changed`, (d) => d.setMeasure(m.table, m, false));
      },
    }),
    tool({
      name: "delete_measure",
      description: "Deletes a measure (visuals and measures using it break).",
      args: z.object({ table: z.string(), name: z.string() }),
      run: async ({ table, name }) => change(`measure [${name}] deleted`, (d) => d.deleteObject(table, "measure", name)),
    }),
    tool({
      name: "create_column",
      description: 'Adds a column. kind "dax": a calculated column (DAX expression, row by row, checked first). For a column computed in Power Query, use a step (step_conditional_column, step_custom_column…).',
      args: z.object({ table: z.string(), name: z.string(), kind: z.enum(["dax"]), expression: z.string(), data_type: z.enum(DATA_TYPES).optional(), check: z.boolean().optional() }),
      run: async ({ table, name, expression, data_type, check }) => {
        await test(`DEFINE COLUMN ${daxRef(table, name)} = ${expression}\nEVALUATE TOPN(3, SELECTCOLUMNS('${table.replace(/'/g, "''")}', "value", ${daxRef(table, name)}))`, check);
        return change(`column [${name}] added`, (d) => d.addCalculatedColumn(table, name, expression, data_type));
      },
    }),
    tool({
      name: "update_column",
      description: "Changes a column's settings (only the given ones): data type, format, hidden, description, sort by another column, default summarization, folder, data category (ImageUrl, WebUrl…), DAX expression (calculated columns only).",
      args: z.object({
        table: z.string(),
        name: z.string(),
        data_type: z.enum(DATA_TYPES).optional(),
        format: z.string().optional(),
        hidden: z.boolean().optional(),
        description: z.string().optional(),
        sort_by: z.string().optional(),
        summarize_by: z.enum(["none", "sum", "average", "min", "max", "count", "distinctCount"]).optional(),
        folder: z.string().optional(),
        category: z.string().optional(),
        expression: z.string().optional(),
      }),
      run: async ({ table, name, data_type, sort_by, summarize_by, ...p }) =>
        change(`column [${name}] changed`, (d) => d.updateColumn(table, name, { ...p, dataType: data_type, sortBy: sort_by, summarizeBy: summarize_by })),
    }),
    tool({
      name: "rename_column",
      description: "Renames a column in the model (its source column keeps its name). DAX using the old name must be updated.",
      args: z.object({ table: z.string(), name: z.string(), new_name: z.string() }),
      run: async ({ table, name, new_name }) => change(`column [${name}] renamed`, (d) => d.renameColumn(table, name, new_name)),
    }),
    tool({
      name: "delete_column",
      description: "Removes a column from the model (to remove it from the data too, add a step_remove_columns instead).",
      args: z.object({ table: z.string(), name: z.string() }),
      run: async ({ table, name }) => change(`column [${name}] deleted`, (d) => d.deleteObject(table, "column", name)),
    }),
    tool({
      name: "create_relationship",
      description: "Relates two tables: from = the many side column ('Sales'[CustomerId]), to = the one side ('Customer'[Id]), whose values must be unique.",
      args: z.object({ from_table: z.string(), from_column: z.string(), to_table: z.string(), to_column: z.string(), both_directions: z.boolean().optional(), active: z.boolean().optional(), one_to_one: z.boolean().optional() }),
      run: async (r) => {
        const state = await open(ref());
        for (const [t, c] of [[r.from_table, r.from_column], [r.to_table, r.to_column]]) {
          const text = state.draft.files.get(state.draft.tablePath(t))!;
          if (!parseTable(text).blocks.some((b) => b.kind === "column" && b.name.toLowerCase() === c.toLowerCase())) throw new Error(`No column "${c}" in '${t}'.`);
        }
        return change(`relationship ${r.from_table} → ${r.to_table}`, (d) =>
          d.addRelationship({ from: { table: r.from_table, column: r.from_column }, to: { table: r.to_table, column: r.to_column }, both: r.both_directions, active: r.active, oneToOne: r.one_to_one }),
        );
      },
    }),
    tool({
      name: "delete_relationship",
      description: "Deletes the relationship between two columns (Table.Column, as listed).",
      args: z.object({ from: z.string().describe("e.g. Sales.CustomerId"), to: z.string().describe("e.g. Customer.Id") }),
      run: async ({ from, to }) => change(`relationship ${from} → ${to} deleted`, (d) => d.deleteRelationship(from, to)),
    }),
    tool({
      name: "get_query",
      description: "The Power Query (M) of a table, step by step.",
      args: z.object({ table: z.string() }),
      readOnly: true,
      run: async ({ table }) => (await open(ref())).draft.query(table).m,
    }),
    step("rename_columns", { renames: z.array(z.object({ from: z.string(), to: z.string() })) }, "renames source columns."),
    step("change_types", { types: z.array(z.object({ column: z.string(), type: mType })) }, "changes column types."),
    step("filter_rows", { column: z.string(), operator: z.enum(["equals", "not_equals", "contains", "not_contains", "greater", "less", "not_empty", "empty"]), value: value.optional() }, "keeps the rows matching a condition."),
    step("replace_values", { column: z.string(), find: z.string(), replace: z.string(), whole: z.boolean().optional().describe("Whole value only (else: inside the text)") }, "replaces values in a column."),
    step("remove_columns", { columns: z.array(z.string()) }, "removes columns."),
    step(
      "conditional_column",
      {
        column: z.string().describe("New column name"),
        rules: z.array(z.object({ if_column: z.string(), operator: z.enum(["equals", "contains", "greater", "less"]), value, then: value })),
        otherwise: value.nullable(),
        type: mType.optional(),
      },
      "adds a column from rules (if … then …, else).",
    ),
    step("custom_column", { column: z.string(), m: z.string().describe('M expression for one row, e.g. [Price] * [Quantity] or Text.Upper([Code])'), type: mType.optional() }, "adds a column computed by a small M expression."),
    step("split_column", { column: z.string(), delimiter: z.string(), into: z.array(z.string()).describe("Names of the new columns") }, "splits a text column by a delimiter."),
    step("transform_text", { columns: z.array(z.string()), transform: z.enum(["trim", "upper", "lower", "proper", "clean"]) }, "trims or changes the case of text columns."),
    step("fill_down", { columns: z.array(z.string()) }, "fills empty cells with the value above."),
    step("remove_duplicates", { columns: z.array(z.string()) }, "keeps one row per value of these columns."),
    tool({
      name: "step_image_base64",
      description: 'Power Query step: a new column with the image of a URL column as "data:image/…;base64,…" (shown as an image by Power BI; the way to show private SharePoint images). The hosts of the URLs are read from the data and written in the query, so Power BI can refresh it. SharePoint / OneDrive images of any size come as a thumbnail (at most 400×400) that fits the 32,766-character limit of Power BI; images of other sites above about 24 KB give no image.',
      args: z.object({ table: z.string(), column: z.string().describe("Column holding the image URLs"), into: z.string().describe("New column name"), name: z.string().optional() }),
      run: async ({ table, column, into, name }) => {
        const urls = await dax(ref(), `EVALUATE DISTINCT(SELECTCOLUMNS(${daxRef(table, column).replace(/\[[^\]]*\]$/, "")}, "u", ${daxRef(table, column)}))`);
        const sources = imageSources(urls.map((r) => String(r.u ?? "")).filter((u) => /^https?:\/\//i.test(u)));
        // The libraries of each SharePoint site (through Graph): the query reads a file through its library's id.
        for (const x of sources.filter((x) => x.sharepoint)) {
          const site = `${x.origin}${encodeURI(x.path)}`;
          const { value } = await graph(`/sites/${new URL(x.origin).host}:${encodeURI(x.path) || "/"}:/drives?$select=id,webUrl`).catch((e) => {
            throw new Error(`Cannot read the libraries of ${site}: ${e instanceof Error ? e.message : e}`);
          });
          x.drives = value.map((d: { id: string; webUrl: string }) => ({ path: decodeURIComponent(new URL(d.webUrl).pathname), id: d.id }));
        }
        return change(`step image_base64 on '${table}'`, (d) => d.addStep(table, { kind: "image_base64", column, into, sources, name }));
      },
    }),
    tool({
      name: "remove_last_step",
      description: "Removes the last Power Query step of a table (its model columns are not changed back).",
      args: z.object({ table: z.string() }),
      run: async ({ table }) => change(`last step of '${table}' removed`, (d) => d.removeLastStep(table)),
    }),
    tool({
      name: "discard_changes",
      description: "Drops your unsaved changes.",
      args: z.object({}),
      run: async () => (discardModelDraft(ref()), "Changes dropped."),
    }),
    tool({
      name: "undo_last_save",
      description: "Puts the model's definition back as it was before your last saved changes (only when the user asks).",
      args: z.object({}),
      run: async () => {
        const files = previous.get(ref().id);
        if (!files) throw new Error("Nothing to undo in this session.");
        await writeModelFiles(ref().id, files, ref().groupId);
        previous.delete(ref().id);
        discardModelDraft(ref());
        return "The model is back to how it was before your last saved changes.";
      },
    }),
    tool({
      name: "preview_table",
      description: "The data of a table as the model holds it now (after the last refresh): its row count, for the given columns how many are blank and distinct, and its first rows (long values shown by their start and length). The way to check that a change or a refresh worked.",
      args: z.object({ table: z.string(), columns: z.array(z.string()).optional().describe("Columns to check and show (default: every column, no counts)"), rows: z.number().optional().describe("First rows shown (default 10, max 50)") }),
      readOnly: true,
      run: async ({ table, columns, rows }) => {
        const t = `'${table.replace(/'/g, "''")}'`;
        const col = (c: string) => `${t}[${c.replace(/]/g, "]]")}]`;
        const label = (s: string) => `"${s.replace(/"/g, '""')}"`;
        const counts = [`"rows", COUNTROWS(${t})`, ...(columns ?? []).flatMap((c) => [`${label(`${c} blank`)}, COUNTBLANK(${col(c)})`, `${label(`${c} distinct`)}, DISTINCTCOUNT(${col(c)})`])];
        const n = Math.min(Math.max(1, Math.round(rows ?? 10)), 50);
        const shown = columns?.length ? `SELECTCOLUMNS(${t}, ${columns.map((c) => `${label(c)}, ${col(c)}`).join(", ")})` : t;
        const [stats, first] = await Promise.all([dax(ref(), `EVALUATE ROW(${counts.join(", ")})`), dax(ref(), `EVALUATE TOPN(${n}, ${shown})`)]);
        const s = stats[0] ?? {};
        const stat = (k: string) => s[`[${k}]`] ?? s[k] ?? 0;
        return [
          `'${table}': ${stat("rows")} rows${(columns ?? []).map((c) => `; [${c}] ${stat(`${c} blank`)} blank, ${stat(`${c} distinct`)} distinct`).join("")}`,
          `First ${n} rows:\n${rowsText(first, n)}`,
        ].join("\n");
      },
    }),
    tool({
      name: "run_dax",
      description: "Runs a DAX query (EVALUATE …) on the saved model, as of its last refresh (max 50 rows shown): check values and totals.",
      args: z.object({ query: z.string() }),
      readOnly: true,
      run: async ({ query }) => rowsText(await dax(ref(), query)),
    }),
  ];
}

export const STEPS_TEXT = `Power Query steps (step_<kind>): ${STEP_KINDS.join(", ")}.`;
