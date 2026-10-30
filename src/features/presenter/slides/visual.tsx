import { z } from "zod";
import { Frame, M, Rich, defineTemplate, fade, hex, plain, pptxTitle, rise, stagger } from "./kit";
import { PowerBiFrame, cropSchema, pptxPowerBi } from "./powerbi";

/** A Power BI page or visual (a zone of its page), with optional points explaining it on the right. */
export const visual = defineTemplate({
  id: "visual",
  description: "A live Power BI report page, or one visual of it (crop = its zone on the 1280×720 page), with optional explaining points on the right.",
  args: z.object({
    title: z.string(),
    report: z.string().describe("Report id or name (a report of the app)"),
    page: z.string().optional().describe("Page name (id) of the report; default: its first page"),
    crop: cropSchema.optional(),
    points: z.array(z.string()).max(5).optional(),
  }),
  Component: ({ args }) => (
    <Frame title={args.title}>
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
  // PowerPoint cannot hold a live report: the picture of the view taken from the screen, linking to it.
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const w = args.points?.length ? 7.4 : 11.7;
    pptxPowerBi(slide, ctx.theme, args, { x: 0.8, y: 2.2, w, h: 4.6 });
    args.points?.forEach((p, i) => slide.addText(plain(p), { x: 8.6, y: 2.3 + i * 0.9, w: 3.9, h: 0.8, fontSize: 14, bullet: true, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text), valign: "top" }));
  },
});
