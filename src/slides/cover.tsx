import { z } from "zod";
import { H, M, Rich, W, brandOf, defineTemplate, ease, hex, plain, reveal, rise, stagger } from "./kit";

/** Cover with a picture: the picture on the left half, the brand, a fine rule, the title, subtitle and description on the right. */
export const cover = defineTemplate({
  id: "cover",
  description: "Cover with a picture: picture (URL) on the left half; on the right the brand name, a fine rule, the title, an italic subtitle and a small description (e.g. a team, a date).",
  kind: "cover",
  bare: true,
  args: z.object({ title: z.string(), subtitle: z.string().optional(), description: z.string().optional(), image: z.string().optional().describe("Picture URL") }),
  Component: ({ args, theme }) => (
    <div className="grid size-full grid-cols-2">
      <M.div variants={reveal} className="overflow-hidden" style={{ background: "color-mix(in srgb, var(--s-secondary) 18%, var(--s-bg))" }}>
        {args.image && <M.img src={args.image} alt="" className="size-full object-cover" variants={{ hidden: { scale: 1.12 }, show: { scale: 1, transition: { duration: 2.6, ease } } }} />}
      </M.div>
      <M.div variants={stagger(0.14, 0.4)} className="flex flex-col justify-end px-[3.4cqw] pb-[7cqw]">
        <M.span variants={rise} className="text-[6.4cqw] leading-none" style={{ fontFamily: "var(--s-brand)" }}>
          {brandOf(theme)}
        </M.span>
        <M.div variants={{ hidden: { scaleX: 0 }, show: { scaleX: 1, transition: { duration: 1.2, ease } } }} className="mt-[1.6cqw] h-px w-full origin-left" style={{ background: "color-mix(in srgb, var(--s-text) 45%, transparent)" }} />
        <div className="mt-[2.4cqw] flex flex-col items-end gap-[0.4cqw] text-right">
          <M.h1 variants={rise} className="text-[2.4cqw] leading-tight font-semibold" style={{ fontFamily: "var(--s-title)" }}>
            <Rich text={args.title} />
          </M.h1>
          {args.subtitle && (
            <M.p variants={rise} className="text-[1.8cqw] leading-tight italic">
              {args.subtitle}
            </M.p>
          )}
          {args.description && (
            <M.p variants={rise} className="text-[1.3cqw] leading-tight italic opacity-80">
              {args.description}
            </M.p>
          )}
        </div>
      </M.div>
    </div>
  ),
  pptx: (slide, args, { theme }) => {
    if (args.image) slide.addImage({ path: args.image, x: 0, y: 0, w: W / 2, h: H, sizing: { type: "cover", w: W / 2, h: H } });
    slide.addText(brandOf(theme), { x: W / 2 + 0.4, y: 3.6, w: W / 2 - 0.8, h: 1.1, fontSize: 54, fontFace: "Bodoni MT", color: hex(theme.text) });
    slide.addShape("line", { x: W / 2 + 0.4, y: 4.8, w: W / 2 - 0.8, h: 0, line: { color: hex(theme.text), width: 0.5, transparency: 45 } });
    slide.addText(plain(args.title), { x: W / 2 + 0.4, y: 5.0, w: W / 2 - 0.8, h: 0.5, fontSize: 20, bold: true, align: "right", fontFace: theme.fontTitle, color: hex(theme.text) });
    if (args.subtitle) slide.addText(args.subtitle, { x: W / 2 + 0.4, y: 5.5, w: W / 2 - 0.8, h: 0.45, fontSize: 16, italic: true, align: "right", fontFace: theme.fontText, color: hex(theme.text) });
    if (args.description) slide.addText(args.description, { x: W / 2 + 0.4, y: 5.95, w: W / 2 - 0.8, h: 0.4, fontSize: 12, italic: true, align: "right", fontFace: theme.fontText, color: hex(theme.text), transparency: 20 });
  },
});
