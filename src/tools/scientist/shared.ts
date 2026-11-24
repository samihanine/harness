/** The analysis document of the screen and its database. */
import { emptyAnalysis, type Analysis, type AnalysisBlock } from "@shared/analysis";
import { engineFor, select } from "@server/data/engine";
import { docOf, setDoc } from "../helpers";
import type { ToolContext } from "../tool";

export const analysisOf = (ctx: ToolContext) => (docOf<Analysis>(ctx) ?? emptyAnalysis()) as Analysis;
export const engine = (ctx: ToolContext) => engineFor(ctx.conversation.scratch, analysisOf(ctx));

export function saveAnalysis(ctx: ToolContext, a: Analysis) {
  setDoc(ctx, { ...a, updatedAt: new Date().toISOString() });
}

/** A block's rows computed again from its SQL (at most 5000 kept). */
export async function computeBlock(ctx: ToolContext, b: AnalysisBlock): Promise<AnalysisBlock> {
  if (b.type === "text") return b;
  const r = select(await engine(ctx), b.sql, 5000);
  return { ...b, rows: r.rows, ...(b.type === "table" ? { columns: r.columns, truncated: r.truncated } : {}) } as AnalysisBlock;
}
