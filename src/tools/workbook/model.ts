/**
 * The Power BI models in the workbook: tables of CUBE formulas (connected: Excel computes them on the model
 * and refreshes them, also in the file downloaded), or a DAX query's rows written as values (a snapshot).
 */
import { z } from "zod";
import { dax } from "@server/pbi/dax";
import { connectionName, ensureConnections, excel, rangePath } from "@server/workbook/excel";
import { tool, type ToolContext } from "../tool";
import { blockAt, colIndex, letters, modelOf, workbookOf } from "./shared";

/** The workbook connected to every model of the screen (once per conversation). */
async function connected(ctx: ToolContext) {
  const id = workbookOf(ctx);
  const models = ctx.scope.datasets ?? [];
  const key = `${id}:${models.map((m) => m.id).sort().join(",")}`;
  if (ctx.conversation.scratch.connected !== key) {
    ctx.status("Connecting the workbook to the models…");
    await ensureConnections(id, models);
    ctx.conversation.scratch.connected = key;
  }
  return id;
}

const parseColumn = (s: string) => {
  const m = s.trim().match(/^'?(.+?)'?\[(.+)\]$/);
  if (!m) throw new Error(`"${s}" is not a column: write 'Table'[Column].`);
  return { table: m[1], column: m[2] };
};
const measureName = (s: string) => s.trim().replace(/^.*?\[(.+)\]$/, "$1");
const mdx = (s: string) => s.replace(/]/g, "]]");
const daxCol = (c: { table: string; column: string }) => `'${c.table.replace(/'/g, "''")}'[${c.column.replace(/]/g, "]]")}]`;
const lit = (v: unknown) => (typeof v === "number" ? String(v) : `"${String(v).replace(/"/g, '""')}"`);
const q = (s: string) => `"${s.replace(/"/g, '""')}"`;

/** A member of a column in the model, as CUBE formulas name it ([T].[C].&[key]). */
function member(c: { table: string; column: string }, value: unknown) {
  const key = typeof value === "string" && /^\d{4}-\d{2}-\d{2}T/.test(value) ? value.slice(0, 19) : String(value);
  return `[${mdx(c.table)}].[${mdx(c.column)}].&[${mdx(key)}]`;
}

export const modelGrid = tool({
  name: "model_grid",
  description:
    "A table connected to a Power BI model, written as CUBE formulas from `at`: one row per value of `rows` (a column, the values with data), and either one column per measure, or — with `columns` — one column per value of that column for the first measure. filters: columns and the values kept (each written as a CUBESET cell above the table, editable by the user). Excel computes it on the model and the downloaded file refreshes it. Use it for figures from the models; combine several models' grids with formulas (e.g. XLOOKUP on the row labels).",
  args: z.object({
    sheet: z.string(),
    at: z.string(),
    model: z.string().optional().describe("Name or id (default: the first model)."),
    rows: z.string().describe("'Table'[Column]"),
    columns: z.string().optional(),
    measures: z.array(z.string()).min(1).describe('["[Revenue]", "[Margin %]"]'),
    filters: z.array(z.object({ column: z.string(), values: z.array(z.union([z.string(), z.number()])) })).optional(),
    max: z.number().int().optional().describe("Most rows (default 200)."),
  }),
  changes: true,
  async run(a, ctx) {
    const id = await connected(ctx);
    const model = modelOf(ctx, a.model);
    const conn = q(connectionName(model));
    const rowCol = parseColumn(a.rows);
    const colCol = a.columns ? parseColumn(a.columns) : undefined;
    const measures = a.measures.map(measureName);
    const filters = (a.filters ?? []).map((f) => ({ col: parseColumn(f.column), values: f.values }));
    const treat = filters.map((f) => `TREATAS({${f.values.map(lit).join(", ")}}, ${daxCol(f.col)})`);
    const ref = { id: model.id, groupId: model.groupId };
    const values = async (c: { table: string; column: string }) =>
      (await dax(ref, `EVALUATE TOPN(${a.max ?? 200}, SUMMARIZECOLUMNS(${[daxCol(c), ...treat, `"m", [${mdx(measures[0])}]`].join(", ")}), ${daxCol(c)}, ASC)`)).map((r) => Object.values(r)[0]);
    const rowValues = await values(rowCol);
    const colValues = colCol ? await values(colCol) : [];
    const m = a.at.replace(/\$/g, "").match(/^([A-Z]+)(\d+)$/i);
    if (!m) throw new Error(`"${a.at}" is not a cell address.`);
    const c0 = colIndex(m[1]);
    let r0 = Number(m[2]);
    const out: (string | number)[][] = [];
    // Filters on top: a label and its CUBESET, referenced by every value.
    const filterRefs: string[] = [];
    for (const f of filters) {
      out.push([`${f.col.column}:`, `=CUBESET(${conn},${q(`{${f.values.map((v) => member(f.col, v)).join(",")}}`)},${q(f.values.join(", "))})`]);
      filterRefs.push(`$${letters(c0 + 1)}$${r0 + out.length - 1}`);
    }
    if (filters.length) out.push([]);
    const head = r0 + out.length;
    const extra = filterRefs.length ? `,${filterRefs.join(",")}` : "";
    if (colCol) {
      out.push([`=CUBEMEMBER(${conn},${q(`[Measures].[${mdx(measures[0])}]`)})`, ...colValues.map((v) => `=CUBEMEMBER(${conn},${q(member(colCol, v))})`)]);
      rowValues.forEach((v, i) => {
        const r = head + 1 + i;
        out.push([`=CUBEMEMBER(${conn},${q(member(rowCol, v))})`, ...colValues.map((_, j) => `=CUBEVALUE(${conn},$${letters(c0)}${r},${letters(c0 + 1 + j)}$${head},$${letters(c0)}$${head}${extra})`)]);
      });
    } else {
      out.push([rowCol.column, ...measures.map((x) => `=CUBEMEMBER(${conn},${q(`[Measures].[${mdx(x)}]`)})`)]);
      rowValues.forEach((v, i) => {
        const r = head + 1 + i;
        out.push([`=CUBEMEMBER(${conn},${q(member(rowCol, v))})`, ...measures.map((_, j) => `=CUBEVALUE(${conn},$${letters(c0)}${r},${letters(c0 + 1 + j)}$${head}${extra})`)]);
      });
    }
    const cols = Math.max(...out.map((r) => r.length));
    const grid = out.map((r) => [...r, ...Array(cols - r.length).fill("")]);
    await excel(id, rangePath(a.sheet, blockAt(a.at, grid.length, cols)), "PATCH", { formulas: grid });
    r0 = head;
    const shown = await excel(id, `${rangePath(a.sheet, blockAt(`${letters(c0)}${head}`, Math.min(6, rowValues.length + 1), cols))}?$select=values`);
    return `Connected table at ${a.sheet}!${blockAt(a.at, grid.length, cols)} (${rowValues.length} rows${colCol ? ` × ${colValues.length} columns` : ""}; header row ${r0}). First rows as Excel computes them: ${JSON.stringify(shown.values)}`;
  },
});

export const loadDax = tool({
  name: "load_dax",
  description: "Write the rows of a DAX query as values from `at` (a snapshot, with a header row), e.g. a detail table to work on with formulas; table: true makes it an Excel table. For figures that must stay connected to the model, use model_grid.",
  args: z.object({ sheet: z.string(), at: z.string(), model: z.string().optional(), query: z.string(), table: z.boolean().optional() }),
  changes: true,
  async run(a, ctx) {
    const id = workbookOf(ctx);
    const model = modelOf(ctx, a.model);
    ctx.status("Querying the model…");
    const rows = await dax({ id: model.id, groupId: model.groupId }, a.query);
    if (!rows.length) return "The query returned no row: nothing written.";
    const columns = Object.keys(rows[0]);
    const values = [columns, ...rows.map((r) => columns.map((c) => (r[c] === null || r[c] === undefined ? "" : (r[c] as string | number))))];
    const CHUNK = Math.max(1, Math.floor(20000 / columns.length));
    const m = a.at.replace(/\$/g, "").match(/^([A-Z]+)(\d+)$/i);
    if (!m) throw new Error(`"${a.at}" is not a cell address.`);
    for (let i = 0; i < values.length; i += CHUNK) await excel(id, rangePath(a.sheet, blockAt(`${m[1]}${Number(m[2]) + i}`, values.slice(i, i + CHUNK).length, columns.length)), "PATCH", { values: values.slice(i, i + CHUNK) });
    const address = blockAt(a.at, values.length, columns.length);
    if (a.table) await excel(id, `/worksheets('${encodeURIComponent(a.sheet)}')/tables/add`, "POST", { address, hasHeaders: true }).catch(() => undefined);
    return `Wrote ${rows.length} rows × ${columns.length} columns at ${a.sheet}!${address}${rows.truncated ? ` (cut at the app's row limit; the query returns more)` : ""}.`;
  },
});
