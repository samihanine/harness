/** A report's structure: pages, visuals with their fields and filters (and all formatting when editable). */
import { z } from "zod";
import { reportStructure as read } from "@server/pbi/report-structure";
import { reportCatalog } from "@server/pbi/catalog";
import { json, tool, type ToolContext } from "../tool";

/** A report of the screen (its report, a reference) or any report the user can open, by id. */
export async function reportRef(ctx: ToolContext, id?: string) {
  const s = ctx.scope;
  if (!id) {
    if (!s.report) throw new Error("Give report_id.");
    return s.report;
  }
  const known = [s.report, ...(s.references ?? [])].find((r) => r && (r.id === id || r.name.toLowerCase() === id.toLowerCase()));
  if (known) return known;
  const any = (await reportCatalog()).find((r) => r.id === id || r.name.toLowerCase() === id.toLowerCase());
  if (!any) throw new Error(`No report "${id}" (list_reports lists them).`);
  return { id: any.id, groupId: any.groupId, name: any.name };
}

export const reportStructure = tool({
  name: "report_structure",
  description:
    "The structure of a report: pages, visuals (type, title, position, fields by role, sort, filters), report and page filters, theme, report measures. detail \"full\" adds every formatting setting (editable reports only). Without report_id: the screen's report.",
  args: z.object({ report_id: z.string().optional(), detail: z.enum(["summary", "full"]).optional(), page: z.string().optional().describe("Only this page (name or display name).") }),
  async run(a, ctx) {
    const ref = await reportRef(ctx, a.report_id);
    const s = await read(ref, a.detail ?? "summary");
    if (a.page) {
      const pages = (s.pages as { name: string; displayName: string }[]).filter((p) => p.name === a.page || p.displayName.toLowerCase() === a.page!.toLowerCase());
      if (!pages.length) throw new Error(`No page "${a.page}". Pages: ${(s.pages as { displayName: string }[]).map((p) => p.displayName).join(", ")}`);
      return json({ ...s, pages });
    }
    return json(s, a.detail === "full" ? 80_000 : 40_000);
  },
});
