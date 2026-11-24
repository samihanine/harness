/** The reports the user can open, to find one by name (e.g. the source of a screenshot in a deck). */
import { z } from "zod";
import { reportCatalog } from "@server/pbi/catalog";
import { tool } from "../tool";

export const listReports = tool({
  name: "list_reports",
  description: "List the Power BI reports the user can open (name, id, workspace), optionally filtered by words. Use it to find a report the user mentions, or the source of a Power BI picture.",
  args: z.object({ search: z.string().optional() }),
  async run(a) {
    const words = (a.search ?? "").toLowerCase().split(/\s+/).filter(Boolean);
    const list = (await reportCatalog()).filter((r) => words.every((w) => `${r.name} ${r.workspace}`.toLowerCase().includes(w)));
    if (!list.length) return a.search ? `No report matches "${a.search}".` : "No report.";
    return list.slice(0, 80).map((r) => `- "${r.name}" id=${r.id}${r.groupId ? ` groupId=${r.groupId}` : ""} (${r.workspace})`).join("\n");
  },
});
