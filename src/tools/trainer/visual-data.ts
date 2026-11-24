/** The data a visual shows now (with the current filters), as the visual summarizes it. */
import { z } from "zod";
import { tool } from "../tool";

export const visualData = tool({
  name: "visual_data",
  description: "The data a visual of the current page shows now, with the current filters (its summarized rows, as CSV). Use it to read exact figures of a visual.",
  args: z.object({ visual: z.string().describe("Name or title of the visual."), rows: z.number().optional().describe("Default 50, max 500.") }),
  async run(a, ctx) {
    const r = (await ctx.browser("report.visualData", { visual: a.visual, rows: Math.min(500, a.rows ?? 50) })) as { title?: string; csv: string };
    return `Visual "${r.title ?? a.visual}":\n${r.csv.trim() || "(no data)"}`;
  },
});
