/** A filter on the report or the page, on any column of the model (when no slicer fits). */
import { z } from "zod";
import { checkValues } from "@server/pbi/dax";
import { tool } from "../tool";

export const setFilter = tool({
  name: "set_filter",
  description:
    "Filter the report (all pages) or the current page on a column, when no slicer of the page is on that column. values: exact values (checked against the model). clear: remove the filter on this column. Prefer set_slicer when a slicer exists for the column.",
  changes: true,
  args: z.object({
    table: z.string(),
    column: z.string(),
    values: z.array(z.union([z.string(), z.number(), z.boolean()])).optional(),
    clear: z.boolean().optional(),
    level: z.enum(["report", "page"]).optional().describe("Default: page."),
  }),
  async run(a, ctx) {
    const report = ctx.scope.report;
    let values = a.values ?? [];
    if (values.length && report?.datasetId) values = (await checkValues({ id: report.datasetId, groupId: report.datasetGroupId }, a.table, a.column, values)) as typeof values;
    await ctx.browser("report.setFilter", { table: a.table, column: a.column, values, clear: a.clear || !values.length, level: a.level ?? "page" });
    return a.clear || !values.length ? `Filter on '${a.table}'[${a.column}] removed.` : `Filter '${a.table}'[${a.column}] in ${JSON.stringify(values)} set on the ${a.level ?? "page"}.`;
  },
});
