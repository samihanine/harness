/** What the user's report shows now: pages, slicers with their selection, filters, visuals. */
import { z } from "zod";
import { tool } from "../tool";
import { viewText, type ViewState } from "./view-text";

export const readView = tool({
  name: "read_view",
  description: "Read the state of the report the user is looking at: pages (current one), slicers (their field and selection), report and page filters, visuals. Read it before and after changing filters.",
  args: z.object({}),
  async run(_, ctx) {
    return viewText((await ctx.browser("report.state")) as ViewState);
  },
});
