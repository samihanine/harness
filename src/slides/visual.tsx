import { z } from "zod";
import { Frame, M, Rich, defineTemplate, fade, hex, plain, pptxTitle, rise, stagger } from "./kit";
import { PowerBiFrame, cropSchema, embedUrl } from "./powerbi";
import { all } from "@/lib/store";

/** A Power BI page or visual (a zone of its page), with optional points explaining it on the right. */
export const visual = defineTemplate({
  id: "visual",
  description: "A live Power BI report page, or one visual of it (crop = its zone on the 1280×720 page), with optional explaining points on the right.",
  args: z.object({
    title: z.string(),
    eyebrow: z.string().optional(),
    report: z.string().describe("Report id or name (a report of the app)"),
    page: z.string().optional().describe("Page name (id) of the report; default: its first page"),
    crop: cropSchema.optional(),
    points: z.array(z.string()).max(5).optional(),
  }),
  Component: ({ args }) => (
    <Frame eyebrow={args.eyebrow} title={args.title}>
      <div className={`grid h-full gap-[3cqw] ${args.points?.length ? "grid-cols-[64fr_36fr]" : "grid-cols-1"}`}>
        <M.div variants={fade} className="min-h-0">
          <PowerBiFrame report={args.report} page={args.page} crop={args.crop} />
        </M.div>
        {args.points?.length ? (
          <M.ul variants={stagger(0.1, 0.4)} className="flex flex-col justify-center gap-[1.4cqw]">
            {args.points.map((p, i) => (
              <M.li key={i} variants={rise} className="flex gap-[1cqw] text-[1.35cqw] leading-snug">
                <span className="mt-[0.55cqw] size-[0.7cqw] shrink-0 rounded-full" style={{ background: "var(--s-primary)" }} />
                <span><Rich text={p} /></span>
              </M.li>
            ))}
          </M.ul>
        ) : null}
      </div>
    </Frame>
  ),
  // PowerPoint cannot hold a live report: a framed placeholder linking to it.
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx, args.eyebrow);
    const w = args.points?.length ? 7.4 : 11.7;
    slide.addShape("roundRect", { x: 0.8, y: 2.2, w, h: 4.6, fill: { color: hex(ctx.theme.primary), transparency: 94 }, line: { color: hex(ctx.theme.primary), transparency: 60, dashType: "dash" }, rectRadius: 0.1 });
    slide.addText([{ text: "Power BI", options: { bold: true, fontSize: 18, breakLine: true } }, { text: "Open the live report", options: { fontSize: 13, hyperlink: { url: reportLinks.get(args.report) ?? "https://app.powerbi.com" } } }], { x: 0.8, y: 2.2, w, h: 4.6, align: "center", valign: "middle", color: hex(ctx.theme.primary), fontFace: ctx.theme.fontText });
    args.points?.forEach((p, i) => slide.addText(plain(p), { x: 8.6, y: 2.3 + i * 0.9, w: 3.9, h: 0.8, fontSize: 14, bullet: true, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text), valign: "top" }));
  },
});

/** Report links for the PowerPoint export (filled before it runs: the export is not a component). */
export const reportLinks = new Map<string, string>();
export async function loadReportLinks() {
  for (const r of await all("reports")) {
    reportLinks.set(r.id, embedUrl(r));
    reportLinks.set(r.name, embedUrl(r));
  }
}
