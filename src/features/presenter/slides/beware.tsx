import { BadgeCheckIcon } from "lucide-react";
import { pptxImage } from "./images";
import { z } from "zod";
import { Frame, M, SlideMarkdown, defineTemplate, ease, fitText, hex, plain, pptxTitle, rise, stagger } from "./kit";

/**
 * Key points to keep in mind: a seal, a heading on a band, then numbered points (markdown) whose text
 * size follows how much there is to say (never overflowing).
 */
export const beware = defineTemplate({
  id: "beware",
  description: "Key points to keep in mind (warnings, rules): a seal, a heading on a band, then 1 to 8 numbered points in markdown (sized to fit).",
  args: z.object({
    title: z.string(),
    heading: z.string().optional().describe("Sentence on the band, e.g. \"Key points to consider when using the report\""),
    items: z.array(z.string()).min(1).max(8).describe("Markdown points (**bold** terms, \"- \" sub-lists)"),
  }),
  Component: ({ args }) => {
    const size = fitText(args.items, 70 * 26, { max: 1.6, min: 0.9 });
    return (
      <Frame title={args.title}>
        <div className="flex h-full flex-col gap-[2cqw]">
          <div className="flex items-center gap-[2.4cqw]">
            <M.span variants={{ hidden: { opacity: 0, rotate: -40, scale: 0.6 }, show: { opacity: 1, rotate: 0, scale: 1, transition: { duration: 1, ease } } }} className="shrink-0" style={{ color: "var(--s-secondary)" }}>
              <BadgeCheckIcon style={{ width: "6.4cqw", height: "6.4cqw" }} strokeWidth={1.1} />
            </M.span>
            {args.heading && (
              <M.div variants={rise} className="flex-1 rounded-[0.3cqw] px-[2cqw] py-[1cqw] text-center text-[1.6cqw]" style={{ background: "color-mix(in srgb, var(--s-text) 6%, var(--s-bg))" }}>
                {args.heading}
              </M.div>
            )}
          </div>
          <M.ol variants={stagger(0.12, 0.4)} className="flex min-h-0 flex-1 flex-col justify-around gap-[1.2cqw] overflow-hidden pl-[4cqw]" style={{ fontSize: `${size}cqw` }}>
            {args.items.map((t, i) => (
              <M.li key={i} variants={rise} className="flex items-center gap-[3cqw] leading-snug">
                <span className="grid size-[3.6cqw] shrink-0 place-items-center rounded-full text-[1.5cqw] text-white" style={{ background: "var(--s-secondary)", clipPath: "polygon(30% 0,70% 0,100% 30%,100% 70%,70% 100%,30% 100%,0 70%,0 30%)" }}>
                  {i + 1}
                </span>
                <div className="min-w-0 flex-1">
                  <SlideMarkdown text={t} />
                </div>
              </M.li>
            ))}
          </M.ol>
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const badge = pptxImage("icon:badge-check", { x: 0.9, y: 2.0, w: 0.9, h: 0.9 });
    if (badge) slide.addImage(badge);
    if (args.heading) slide.addText(args.heading, { x: 2.2, y: 2.1, w: 10.3, h: 0.6, fontSize: 18, align: "center", fill: { color: "F3F4F6" }, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    const h = Math.min(1, 4.1 / args.items.length);
    args.items.forEach((t, i) => {
      const y = 2.95 + i * h;
      slide.addShape("octagon", { x: 1.3, y, w: 0.5, h: 0.5, fill: { color: hex(ctx.theme.secondary ?? "#8b5a2b") }, line: { type: "none" } as never });
      slide.addText(String(i + 1), { x: 1.3, y, w: 0.5, h: 0.5, fontSize: 14, color: "FFFFFF", align: "center", valign: "middle" });
      slide.addText(plain(t).replace(/\n+/g, " "), { x: 2.2, y: y - 0.1, w: 10.3, h: h, fontSize: 14, valign: "middle", fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    });
  },
});
