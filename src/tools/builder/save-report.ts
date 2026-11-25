/** Writes the draft to the report (checked first: overlaps, visuals outside the page), and reloads it on screen. */
import { z } from "zod";
import { pagesOf, saveDraft, visualsOf } from "@server/pbi/draft";
import { layoutProblems } from "@server/pbi/layout";
import { forgetReport } from "@server/capture/browser";
import { tool } from "../tool";
import { draft } from "./shared";

export const saveReport = tool({
  name: "save_report",
  description: "Save the draft to the report (the user then sees it). Checks the layout first: problems are returned instead of saving, unless force is true. Then look_at_report to check the result visually.",
  args: z.object({ force: z.boolean().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    if (!d.changed.size) return "Nothing to save.";
    // Only the pages changed in this draft are checked (what was already on the others is the user's).
    const touched = (name: string) => [...d.changed].some((c) => c === `definition/pages/${name}` || c.startsWith(`definition/pages/${name}/`));
    const problems = pagesOf(d).filter((p) => touched(p.name)).flatMap((p) => layoutProblems(visualsOf(d, p.name).filter((v) => v.json.position && !v.json.parentGroupName).map((v) => ({ name: v.json.name, ...v.json.position, hidden: v.json.isHidden, type: v.json.visual?.visualType })), { width: p.width ?? 1280, height: p.height ?? 720 }).map((x) => `"${p.displayName}": ${x}`));
    if (problems.length && !a.force) throw new Error(`Not saved, the layout has problems (fix them with layout_page, or save with force when intended):\n${problems.join("\n")}`);
    ctx.status("Saving the report…");
    const n = await saveDraft(d);
    ctx.effect("reload");
    await forgetReport(d.ref.id);
    // The service serves the new definition a little after it is written.
    await new Promise((r) => setTimeout(r, 2500));
    return `Saved (${n} change(s)). The report on screen reloads.`;
  },
});

export const discardChanges = tool({
  name: "discard_changes",
  description: "Forget the unsaved changes of the draft (the report is read again).",
  args: z.object({}),
  async run(_, ctx) {
    delete ctx.conversation.scratch.draft;
    return "Unsaved changes discarded.";
  },
});
