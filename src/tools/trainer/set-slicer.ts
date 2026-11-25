/** Changes a slicer of the user's report: values (checked against the model), a range, or nothing selected. */
import { z } from "zod";
import { checkValues } from "@server/pbi/dax";
import { tool } from "../tool";
import { viewText, type ViewState } from "./view-text";

export const setSlicer = tool({
  name: "set_slicer",
  description:
    "Select values in a slicer of the current page (by its name or title; find the right slicer by the column it is on, even when its title differs from the user's words). values: exact values of the column (matched without case; a wrong value is refused with the valid ones). range: for numbers / dates (ISO). clear: select everything again.",
  args: z.object({
    slicer: z.string(),
    values: z.array(z.union([z.string(), z.number(), z.boolean()])).optional(),
    range: z.object({ from: z.union([z.string(), z.number()]).optional(), to: z.union([z.string(), z.number()]).optional() }).optional(),
    clear: z.boolean().optional(),
  }),
  changes: true,
  async run(a, ctx) {
    const view = (await ctx.browser("report.state")) as ViewState;
    const s = view.slicers.find((x) => x.name === a.slicer) ?? view.slicers.find((x) => (x.title ?? "").toLowerCase() === a.slicer.toLowerCase());
    if (!s) throw new Error(`No slicer "${a.slicer}" on the current page. Slicers: ${view.slicers.map((x) => `"${x.title ?? ""}" (${x.name}) on ${x.fields.map((f) => `${f.table}[${f.column}]`).join(">")}`).join(", ") || "none"}. A slicer may be on another page: go_to_page first, or use set_filter.`);
    let values = a.values;
    const field = s.fields[0];
    const report = ctx.scope.report;
    if (values?.length && field && report?.datasetId) values = (await checkValues({ id: report.datasetId, groupId: report.datasetGroupId }, field.table, field.column, values)) as typeof values;
    const done = (await ctx.browser("report.setSlicer", { name: s.name, values, range: a.range, clear: a.clear || (!values?.length && !a.range) })) as { note?: string };
    const after = (await ctx.browser("report.state")) as ViewState;
    return `Done.${done?.note ? ` ${done.note}` : ""} ${viewText(after).split("\n").filter((l) => l.includes(`name=${s.name}`) || l.startsWith("Page filters")).join("\n")}`;
  },
});
