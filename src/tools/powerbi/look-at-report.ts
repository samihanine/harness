/**
 * A picture of a report as the user sees it (or of any page, with filters): taken by a headless browser
 * on the server (no screen sharing), with the list of what is on the page.
 */
import { z } from "zod";
import { captureReport } from "@server/capture/browser";
import { reportRef } from "./report-structure";
import { base64, tool } from "../tool";

export const filterSpec = z.array(z.object({ table: z.string(), column: z.string(), values: z.array(z.union([z.string(), z.number(), z.boolean()])) }));
export const toFilters = (f?: z.infer<typeof filterSpec>) =>
  (f ?? []).map((x) => ({ $schema: "http://powerbi.com/product/schema#basic", target: { table: x.table, column: x.column }, operator: "In", values: x.values, filterType: 1 }));

export const lookAtReport = tool({
  name: "look_at_report",
  description:
    "See a report page as a picture. Without arguments: exactly what the user sees now on this screen (page, filters, slicers). Or a page (and optionally one visual) of any report, with filters. Use it to check a visual result, read a chart, or compare with a picture.",
  args: z.object({
    report_id: z.string().optional(),
    page: z.string().optional().describe("Page name or display name (default: the user's current view)."),
    visual: z.string().optional().describe("Keep only this visual (name or title)."),
    filters: filterSpec.optional(),
  }),
  async run(a, ctx) {
    const ref = await reportRef(ctx, a.report_id);
    const onScreen = !a.report_id || a.report_id === ctx.scope.report?.id;
    // The user's view: its state taken from the embedded report (when this screen shows it).
    const state = onScreen && !a.page && !a.filters ? await ctx.browser("report.bookmark").catch(() => undefined) : undefined;
    ctx.status("Taking a picture of the report…");
    const shot = await captureReport(ref, { state: state ?? undefined, page: a.page, visual: a.visual, filters: toFilters(a.filters) });
    return {
      text: `Report "${ref.name}", page "${shot.page.displayName}"${a.visual ? `, visual ${a.visual}` : ""}${state ? " (the user's current view)" : ""}. Visuals on the page: ${shot.visuals.map((v) => `${v.type} "${v.title ?? ""}" (${v.name})`).join(", ")}. Pages: ${shot.pages.map((p) => p.displayName).join(", ")}.`,
      images: [{ mediaType: "image/png", data: base64(shot.png) }],
    };
  },
});
