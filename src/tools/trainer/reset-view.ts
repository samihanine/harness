/** Back to the report as it opened: every slicer and filter the AI or the user changed undone. */
import { z } from "zod";
import { tool } from "../tool";

export const resetView = tool({
  name: "reset_view",
  description: "Undo every filter and slicer change: the report as it opened.",
  changes: true,
  args: z.object({}),
  async run(_, ctx) {
    await ctx.browser("report.reset");
    return "The report is back to its first state.";
  },
});
