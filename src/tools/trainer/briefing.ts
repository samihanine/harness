/**
 * What the Trainer's AI is told when a report opens: each slicer of the report's pages with the column
 * it is on and examples of its values, so it finds the right slicer for the user's words even when the
 * slicer has no title (read from the browser's embed, values from the model).
 */
import { columnValues } from "@server/pbi/dax";
import type { ToolContext } from "../tool";

type SlicerInfo = { page: string; name: string; title?: string; fields: { table: string; column: string }[]; selected: string };

export async function slicerBriefing(ctx: ToolContext) {
  const list = (await ctx.browser("report.slicers").catch(() => undefined)) as SlicerInfo[] | undefined;
  if (!list?.length) return undefined;
  const r = ctx.scope.report;
  const lines = await Promise.all(
    list.map(async (s) => {
      const f = s.fields[0];
      const values = f && r?.datasetId ? await columnValues({ id: r.datasetId, groupId: r.datasetGroupId }, f.table, f.column).catch(() => undefined) : undefined;
      const sample = values?.values.length ? ` — values e.g. ${values.values.slice(0, 12).map((v) => JSON.stringify(v)).join(", ")}${values.more || values.values.length > 12 ? ", …" : ""}` : "";
      return `- page "${s.page}": slicer "${s.title || "(no title)"}" name=${s.name} on ${s.fields.map((x) => `'${x.table}'[${x.column}]`).join(" > ") || "?"}${sample}; now ${s.selected}`;
    }),
  );
  return `Slicers of this report (find the one to use by its column and values):\n${lines.join("\n")}`;
}
