import { z } from "zod";
import { Frame, M, Rich, defineTemplate, fade, hex, plain, pptxTitle, rise } from "./kit";
import { PowerBiFrame, pptxPowerBi } from "./powerbi";

/** A whole Power BI report page, live, on the full width (an optional sentence above it). */
export const report = defineTemplate({
  id: "report",
  description: "A whole live Power BI report page on the full width, with an optional lead sentence. report: the report's id; page: the page's name (id) from the context.",
  args: z.object({ title: z.string(), lead: z.string().optional(), report: z.string().describe("Report id (or name)"), page: z.string().optional().describe("Page name (id); default: its first page") }),
  Component: ({ args }) => (
    <Frame title={args.title}>
      <div className="flex h-full flex-col gap-[1.2cqw]">
        {args.lead && (
          <M.p variants={rise} className="max-w-[70cqw] text-[1.4cqw] leading-snug opacity-80">
            <Rich text={args.lead} />
          </M.p>
        )}
        <M.div variants={fade} className="min-h-0 flex-1">
          <PowerBiFrame report={args.report} page={args.page} />
        </M.div>
      </div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    if (args.lead) slide.addText(plain(args.lead), { x: 0.8, y: 2.1, w: 11.5, h: 0.5, fontSize: 15, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    pptxPowerBi(slide, ctx.theme, args, { x: 0.8, y: 2.7, w: 11.7, h: 4.2 });
  },
});
