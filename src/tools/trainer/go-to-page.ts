/** Shows another page of the user's report. */
import { z } from "zod";
import { tool } from "../tool";
import { viewText, type ViewState } from "./view-text";

export const goToPage = tool({
  name: "go_to_page",
  description: "Show another page of the report (name or display name).",
  changes: true,
  args: z.object({ page: z.string() }),
  async run(a, ctx) {
    await ctx.browser("report.setPage", { page: a.page });
    return viewText((await ctx.browser("report.state")) as ViewState);
  },
});
