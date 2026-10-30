import { z } from "zod";
import { pptxImage } from "./images";
import { Frame, M, SlideMarkdown, defineTemplate, fade, hex, plain, pptxTitle, rise } from "./kit";

/** Markdown text on the left, a picture on the right shown whole (contain: never cropped). */
export const image = defineTemplate({
  id: "image",
  description: "A picture shown whole on the right part (never cropped: a screenshot, a diagram, a photo) with markdown text on the left (paragraphs, lists, **bold**). image: a library image \"img:<id>\" (or a URL); caption: a short line under it.",
  args: z.object({ title: z.string(), text: z.string(), image: z.string().describe('Picture: "img:<id>" from the library, or a URL'), caption: z.string().optional() }),
  Component: ({ args }) => (
    <Frame title={args.title}>
      <div className="grid h-full grid-cols-[42fr_58fr] gap-[3cqw]">
        <M.div variants={rise} className="flex min-h-0 flex-col justify-center overflow-hidden text-[1.45cqw] leading-relaxed">
          <SlideMarkdown text={args.text} />
        </M.div>
        <M.figure variants={fade} className="flex min-h-0 flex-col items-center justify-center gap-[0.8cqw]">
          {args.image ? <img src={args.image} alt={args.caption ?? ""} className="min-h-0 w-full flex-1 object-contain" /> : <div className="flex w-full flex-1 items-center justify-center rounded-[0.6cqw] text-[1.2cqw] opacity-50" style={{ background: "color-mix(in srgb, var(--s-text) 4%, var(--s-bg))" }}>Image not found</div>}
          {args.caption && <figcaption className="text-center text-[1.05cqw] opacity-60">{args.caption}</figcaption>}
        </M.figure>
      </div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const lines = args.text.split("\n").filter((l) => l.trim());
    slide.addText(
      lines.map((l) => {
        const bullet = /^\s*[-*]\s+/.test(l);
        return { text: plain(l.replace(/^\s*[-*]\s+/, "").replace(/^#+\s*/, "")), options: { bullet, breakLine: true, bold: /^#/.test(l) } };
      }),
      { x: 0.8, y: 2.2, w: 4.9, h: 4.6, fontSize: 15, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text), valign: "middle", paraSpaceAfter: 6 },
    );
    const box = { x: 6.1, y: 2.0, w: 6.6, h: args.caption ? 4.5 : 5.0 };
    const image = args.image && pptxImage(args.image, box);
    if (image) slide.addImage(image);
    if (args.caption) slide.addText(plain(args.caption), { x: box.x, y: box.y + box.h + 0.05, w: box.w, h: 0.4, fontSize: 11, align: "center", color: hex(ctx.theme.text), transparency: 40, fontFace: ctx.theme.fontText });
  },
});
