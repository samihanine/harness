/** What the analysis shows: text, key figures, tables, charts and pivots, each computed from SQL. */
import { z } from "zod";
import type { AnalysisBlock } from "@shared/analysis";
import { rowsText } from "@server/data/engine";
import { tool } from "../tool";
import { analysisOf, computeBlock, saveAnalysis } from "./shared";

const block = z.discriminatedUnion("type", [
  z.object({ type: z.literal("text"), markdown: z.string() }),
  z.object({ type: z.literal("kpis"), title: z.string().optional(), sql: z.string().describe("One row; each item names a column of it."), items: z.array(z.object({ label: z.string(), column: z.string(), format: z.string().optional().describe('"number", "percent", "currency:EUR", "compact"') })) }),
  z.object({ type: z.literal("table"), title: z.string().optional(), sql: z.string() }),
  z.object({ type: z.literal("chart"), title: z.string().optional(), sql: z.string(), chart: z.object({ kind: z.enum(["bar", "column", "line", "area", "pie", "scatter"]), x: z.string(), y: z.array(z.string()), stacked: z.boolean().optional() }) }),
  z.object({ type: z.literal("pivot"), title: z.string().optional(), sql: z.string(), pivot: z.object({ rows: z.array(z.string()), columns: z.string().optional(), value: z.string(), aggregate: z.enum(["sum", "avg", "count", "min", "max"]) }) }),
]);

export const setBlock = tool({
  name: "set_block",
  description:
    "Add or replace a block of the analysis shown to the user: text (markdown: the findings, the method), kpis (key figures from a one-row query), table, chart (bar, column, line, area, pie, scatter: x column, y columns), pivot (rows, columns, value, aggregate). Each block's SQL runs on the analysis tables and steps. id: the block to replace (else a new one); position: where to insert (0 = first).",
  args: z.object({ id: z.string().optional(), position: z.number().optional(), block }),
  changes: true,
  async run(a, ctx) {
    const an = analysisOf(ctx);
    const id = a.id ?? `b${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`;
    const computed = await computeBlock(ctx, { ...(a.block as object), id } as AnalysisBlock);
    const blocks = an.blocks.filter((b) => b.id !== id);
    const at = a.id ? an.blocks.findIndex((b) => b.id === id) : (a.position ?? blocks.length);
    blocks.splice(at < 0 ? blocks.length : at, 0, computed);
    saveAnalysis(ctx, { ...an, blocks });
    const rows = "rows" in computed && computed.rows ? computed.rows : undefined;
    return `Block ${id} (${computed.type}) ${a.id ? "replaced" : "added"}.${rows ? `\n${rowsText(Object.keys(rows[0] ?? {}), rows, 15)}` : ""}`;
  },
});

export const removeBlock = tool({
  name: "remove_block",
  description: "Remove a block of the analysis.",
  args: z.object({ id: z.string() }),
  changes: true,
  async run(a, ctx) {
    const an = analysisOf(ctx);
    saveAnalysis(ctx, { ...an, blocks: an.blocks.filter((b) => b.id !== a.id) });
    return `Block ${a.id} removed.`;
  },
});

export const readAnalysis = tool({
  name: "read_analysis",
  description: "The analysis as the user sees it: title, sources, steps, blocks (id, type, title, SQL, first rows).",
  args: z.object({}),
  async run(_, ctx) {
    const an = analysisOf(ctx);
    return [
      `Title: ${an.title}`,
      `Sources: ${an.sources.map((s) => `${s.name} (${s.kind === "dax" ? `DAX on ${s.datasetName}` : `file ${s.fileName}${s.sheet ? `/${s.sheet}` : ""}`}, ${s.rows ?? "?"} rows)`).join("; ") || "none"}`,
      `Steps: ${an.steps.map((s) => `${s.name} = ${s.sql.replace(/\s+/g, " ").slice(0, 200)}`).join("\n") || "none"}`,
      `Blocks:\n${an.blocks.map((b) => `- ${b.id} ${b.type}${"title" in b && b.title ? ` "${b.title}"` : ""}${b.type === "text" ? `: ${b.markdown.slice(0, 200)}` : `: ${b.sql.replace(/\s+/g, " ").slice(0, 200)}`}`).join("\n") || "none"}`,
    ].join("\n");
  },
});

export const setTitle = tool({
  name: "set_title",
  description: "Rename the analysis.",
  args: z.object({ title: z.string() }),
  changes: true,
  async run(a, ctx) {
    saveAnalysis(ctx, { ...analysisOf(ctx), title: a.title });
    return "Renamed.";
  },
});
