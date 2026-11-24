/** The report being built: pages and visuals of the draft (unsaved changes included). */
import { z } from "zod";
import { reportStructureFromParts } from "@server/pbi/report-structure";
import { json, tool } from "../tool";
import { draft } from "./shared";

export const reportOverview = tool({
  name: "report_overview",
  description: "The report being built, as the draft holds it (unsaved changes included): pages (name, size), visuals (type, title, position, fields, filters), report filters, theme, report measures. Read it first, and after changes.",
  args: z.object({ page: z.string().optional() }),
  async run(a, ctx) {
    const d = await draft(ctx);
    const s = await reportStructureFromParts(d.parts, "summary");
    const pages = a.page ? s.pages.filter((p: any) => p.name === a.page || p.displayName.toLowerCase() === a.page!.toLowerCase()) : s.pages;
    return `${d.changed.size ? `Unsaved changes: ${d.changed.size} file(s) (save_report writes them).` : "No unsaved change."}\n${json({ ...s, pages })}`;
  },
});
