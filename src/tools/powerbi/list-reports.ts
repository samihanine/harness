/**
 * The reports of the team's catalog (Confluence: name, what each is for), to pick the right one for a question,
 * find a report the user mentions, or the source of a Power BI picture.
 */
import { z } from "zod";
import { library } from "@server/library";
import { tool } from "../tool";

export const listReports = tool({
  name: "list_reports",
  description: "The reports of the team's catalog: name, id and what each is for (its description), optionally filtered by words. Use it to find the report that answers a question best, a report the user mentions, or the source of a Power BI picture.",
  args: z.object({ search: z.string().optional() }),
  async run(a) {
    const words = (a.search ?? "").toLowerCase().split(/\s+/).filter(Boolean);
    const list = (await library("report")).filter((r) => r.id && words.every((w) => `${r.name} ${r.description}`.toLowerCase().includes(w)));
    if (!list.length) return a.search ? `No report of the catalog matches "${a.search}".` : "The catalog has no report.";
    return list.slice(0, 80).map((r) => `- "${r.name}" id=${r.id}${r.groupId ? ` groupId=${r.groupId}` : ""}${r.appId ? ` appId=${r.appId}` : ""}${r.description ? `: ${r.description.replace(/\s+/g, " ").slice(0, 400)}` : ""}`).join("\n");
  },
});
