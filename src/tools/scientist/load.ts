/** Sources loaded as SQL tables: a DAX query's result, or a sheet of an uploaded file. */
import { z } from "zod";
import type { Source } from "@shared/analysis";
import { loadSource, rowsText, safeName, select, ident } from "@server/data/engine";
import { modelOf } from "../helpers";
import { tool } from "../tool";
import { analysisOf, engine, saveAnalysis } from "./shared";

async function load(ctx: Parameters<typeof saveAnalysis>[0], source: Source) {
  const e = await engine(ctx);
  const cols = await loadSource(e, source);
  const rows = e.loaded.get(source.name)?.rows ?? 0;
  const a = analysisOf(ctx);
  saveAnalysis(ctx, { ...a, sources: [...a.sources.filter((s) => s.name !== source.name), { ...source, rows }] });
  const sample = select(e, `SELECT * FROM ${ident(source.name)} LIMIT 5`);
  return `Table ${source.name}: ${rows} rows, columns ${cols.map((c) => `${c.name} ${c.type.toLowerCase()}`).join(", ")}\n${rowsText(sample.columns, sample.rows)}`;
}

export const loadDax = tool({
  name: "load_dax",
  description:
    "Load the result of a DAX query as a SQL table of the analysis (e.g. revenue by month and region from a model). Keep it at the grain you need (SUMMARIZECOLUMNS with the columns and measures to analyse); column names lose their table prefix.",
  args: z.object({ name: z.string().describe("SQL table name, e.g. sales_by_month."), query: z.string(), dataset_id: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = modelOf(ctx, a.dataset_id);
    return load(ctx, { kind: "dax", name: safeName(a.name), datasetId: d.id, groupId: d.groupId, datasetName: d.name, query: a.query });
  },
});

export const loadFile = tool({
  name: "load_file",
  description: "Load a sheet of an uploaded Excel / CSV file as a SQL table of the analysis (its first filled row gives the column names).",
  args: z.object({ name: z.string(), file_id: z.string(), sheet: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const f = (ctx.scope.files ?? []).find((x) => x.id === a.file_id || x.name === a.file_id);
    if (!f) throw new Error(`No file "${a.file_id}". Files: ${(ctx.scope.files ?? []).map((x) => `"${x.name}" (${x.id})`).join(", ") || "none: the user can add some in the Sources tab"}`);
    return load(ctx, { kind: "file", name: safeName(a.name), fileId: f.id, fileName: f.name, sheet: a.sheet });
  },
});
