/** A picture of a Power BI page or visual, kept as a file for a slide (markdown `![](file:<id>)` or a PowerPoint picture). */
import { z } from "zod";
import { captureReport } from "@server/capture/browser";
import { saveFile } from "@server/files";
import { filterSpec, toFilters } from "../powerbi/look-at-report";
import { reportRef } from "../powerbi/report-structure";
import { base64, tool, type ToolContext } from "../tool";

export const captureArgs = { report_id: z.string(), page: z.string().optional(), visual: z.string().optional(), filters: filterSpec.optional() };

export async function capture(ctx: ToolContext, a: { report_id: string; page?: string; visual?: string; filters?: z.infer<typeof filterSpec> }) {
  const ref = await reportRef(ctx, a.report_id);
  ctx.status("Taking a picture of the report…");
  const shot = await captureReport(ref, { page: a.page, visual: a.visual, filters: toFilters(a.filters), width: 1600, height: 900 });
  const file = await saveFile(shot.png, `${ref.name} - ${shot.page.displayName}${a.visual ? ` - ${a.visual}` : ""}.png`, "image/png");
  return { file, shot, ref };
}

export const capturePowerbi = tool({
  name: "capture_powerbi",
  description: "Take a picture of a Power BI report page or of one visual (with filters) for a slide. Returns its source `file:<id>` to use as a picture (markdown ![caption](file:<id>), or replace_picture). You see the picture to check it.",
  args: z.object(captureArgs),
  async run(a, ctx) {
    const { file, shot, ref } = await capture(ctx, a);
    return { text: `Picture of "${ref.name}" / "${shot.page.displayName}"${a.visual ? ` / ${a.visual}` : ""}: file:${file.id}`, images: [{ mediaType: "image/png", data: base64(shot.png) }] };
  },
});
