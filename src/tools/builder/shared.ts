/** The draft of the screen's report, for the Builder's tools. */
import { draftOf } from "@server/pbi/draft";
import { reportOf } from "../helpers";
import type { ToolContext } from "../tool";

export async function draft(ctx: ToolContext) {
  const r = reportOf(ctx);
  if (r.access === "view") throw new Error("This report is view only: its definition cannot be changed. Choose a report you can edit, or create one.");
  return draftOf(ctx.conversation.scratch, r);
}
