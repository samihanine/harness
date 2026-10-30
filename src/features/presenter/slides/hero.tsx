import { z } from "zod";
import { pptxImage } from "./images";
import { Eyebrow, H, M, Rich, Rule, W, defineTemplate, ease, hex, plain, reveal, rise, stagger } from "./kit";

/** The cover: big title, subtitle, author and date; an optional picture on the right half. */
export const hero = defineTemplate({
  id: "hero",
  description: "Cover of the deck: title, subtitle, author/date line, optional picture (URL).",
  kind: "cover",
  bare: true,
  args: z.object({ title: z.string(), subtitle: z.string().optional(), eyebrow: z.string().optional(), footer: z.string().optional().describe("e.g. author · date"), image: z.string().optional() }),
  Component: ({ args }) => (
    <div className={`grid size-full ${args.image ? "grid-cols-[52fr_48fr]" : "grid-cols-1"}`}>
      <M.div variants={stagger(0.12, 0.2)} className="flex flex-col justify-center gap-[1.6cqw] px-[7cqw]">
        {args.eyebrow && <Eyebrow>{args.eyebrow}</Eyebrow>}
        <M.h1 variants={rise} className="max-w-[44cqw] text-[5.4cqw] leading-[1.02] font-bold tracking-tight" style={{ fontFamily: "var(--s-title)" }}>
          <Rich text={args.title} />
        </M.h1>
        <Rule />
        {args.subtitle && (
          <M.p variants={rise} className="max-w-[40cqw] text-[1.7cqw] leading-snug opacity-75">
            <Rich text={args.subtitle} />
          </M.p>
        )}
        {args.footer && <M.p variants={rise} className="mt-[3cqw] text-[1.1cqw] font-medium opacity-60">{args.footer}</M.p>}
      </M.div>
      {args.image && (
        <M.div variants={reveal} className="overflow-hidden">
          <M.img src={args.image} alt="" className="size-full object-cover" variants={{ hidden: { scale: 1.15 }, show: { scale: 1, transition: { duration: 2.4, ease } } }} />
        </M.div>
      )}
    </div>
  ),
  pptx: (slide, args, { theme }) => {
    const w = args.image ? W * 0.52 : W;
    const image = args.image && pptxImage(args.image, { x: w, y: 0, w: W - w, h: H }, "cover");
    if (image) slide.addImage(image);
    if (args.eyebrow) slide.addText(args.eyebrow.toUpperCase(), { x: 0.9, y: 2.1, w: w - 1.5, h: 0.4, fontSize: 12, bold: true, color: hex(theme.primary), charSpacing: 3, fontFace: theme.fontText });
    slide.addText(plain(args.title), { x: 0.9, y: 2.5, w: w - 1.5, h: 1.8, fontSize: 44, bold: true, fontFace: theme.fontTitle, color: hex(theme.text), valign: "top" });
    slide.addShape("rect", { x: 0.9, y: 4.4, w: 0.8, h: 0.05, fill: { color: hex(theme.primary) }, line: { color: hex(theme.primary) } });
    if (args.subtitle) slide.addText(plain(args.subtitle), { x: 0.9, y: 4.6, w: w - 1.5, h: 0.9, fontSize: 18, fontFace: theme.fontText, color: hex(theme.text), transparency: 25 });
    if (args.footer) slide.addText(args.footer, { x: 0.9, y: 6.3, w: w - 1.5, h: 0.4, fontSize: 12, fontFace: theme.fontText, color: hex(theme.text), transparency: 40 });
  },
});
