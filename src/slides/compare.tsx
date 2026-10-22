import { z } from "zod";
import { Frame, M, defineTemplate, fade, hex, pptxTitle, stagger } from "./kit";
import { PowerBiFrame, cropSchema } from "./powerbi";
import { reportLinks } from "./visual";

const side = z.object({ label: z.string(), report: z.string().describe("Report id or name"), page: z.string().optional(), crop: cropSchema.optional() });

/** Two Power BI views side by side (before / after, two pages, two visuals). */
export const compare = defineTemplate({
  id: "compare",
  description: "Two Power BI views side by side (before/after, two pages or visuals), each with a label.",
  args: z.object({ title: z.string(), eyebrow: z.string().optional(), left: side, right: side }),
  Component: ({ args }) => (
    <Frame eyebrow={args.eyebrow} title={args.title}>
      <M.div variants={stagger(0.2, 0.3)} className="grid h-full grid-cols-2 gap-[2cqw]">
        {[args.left, args.right].map((s, i) => (
          <M.div key={i} variants={fade} className="flex min-h-0 flex-col gap-[0.8cqw]">
            <span className="text-[1.2cqw] font-semibold tracking-wide uppercase" style={{ color: i ? "var(--s-primary)" : undefined, opacity: i ? 1 : 0.6 }}>{s.label}</span>
            <div className="min-h-0 flex-1">
              <PowerBiFrame report={s.report} page={s.page} crop={s.crop} />
            </div>
          </M.div>
        ))}
      </M.div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx, args.eyebrow);
    [args.left, args.right].forEach((s, i) => {
      const x = 0.8 + i * 6.0;
      slide.addText(s.label.toUpperCase(), { x, y: 2.1, w: 5.7, h: 0.4, fontSize: 12, bold: true, color: hex(i ? ctx.theme.primary : ctx.theme.text), fontFace: ctx.theme.fontText });
      slide.addShape("roundRect", { x, y: 2.6, w: 5.7, h: 4.2, fill: { color: hex(ctx.theme.primary), transparency: 94 }, line: { color: hex(ctx.theme.primary), transparency: 60, dashType: "dash" }, rectRadius: 0.1 });
      slide.addText("Open the live report", { x, y: 2.6, w: 5.7, h: 4.2, align: "center", fontSize: 13, color: hex(ctx.theme.primary), hyperlink: { url: reportLinks.get(s.report) ?? "https://app.powerbi.com" }, fontFace: ctx.theme.fontText });
    });
  },
});
