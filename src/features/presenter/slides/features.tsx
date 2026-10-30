import { icons, SparklesIcon } from "lucide-react";
import { z } from "zod";
import { pptxImage } from "./images";
import { Frame, M, SlideMarkdown, defineTemplate, fitText, hex, plain, pptxTitle, rise, stagger } from "./kit";

/** A lucide icon from its name ("store", "users-round", "globe"…), a sparkle when unknown. */
const iconOf = (name: string) => icons[name.replace(/(^|[-_ ])(\w)/g, (_, __, c: string) => c.toUpperCase()) as keyof typeof icons] ?? SparklesIcon;

/** Up to three feature cards side by side: a title, a markdown text, an icon in the corner. */
export const features = defineTemplate({
  id: "features",
  description: "1 to 3 feature cards side by side (laid out automatically): each a title, a markdown text (bold, \"- \" lists) and a lucide icon name (e.g. store, users-round, globe, chart-line).",
  args: z.object({
    title: z.string(),
    items: z.array(z.object({ title: z.string(), text: z.string(), icon: z.string().optional() })).min(1).max(3),
  }),
  Component: ({ args }) => {
    const size = fitText([args.items.reduce((a, b) => (a.text.length > b.text.length ? a : b)).text], (84 / args.items.length) * 26, { max: 1.5, min: 0.95 });
    return (
      <Frame title={args.title}>
        <M.div variants={stagger(0.15, 0.2)} className="grid h-full gap-[2.6cqw]" style={{ gridTemplateColumns: `repeat(${args.items.length}, minmax(0, 1fr))`, padding: args.items.length < 3 ? "0 6cqw" : undefined }}>
          {args.items.map((it, i) => {
            const Icon = iconOf(it.icon ?? "");
            return (
              <M.div key={i} variants={rise} className="relative flex min-h-0 flex-col gap-[1.6cqw] overflow-hidden px-[2cqw] pt-[2cqw] pb-[5cqw] shadow-[0_0.3cqw_1cqw_rgba(0,0,0,0.08)]" style={{ background: "color-mix(in srgb, var(--s-secondary) 14%, var(--s-bg))" }}>
                <span className="text-center text-[1.6cqw] font-bold">{it.title}</span>
                <div className="min-h-0 leading-[1.8]" style={{ fontSize: `${size}cqw` }}>
                  <SlideMarkdown text={it.text} />
                </div>
                <Icon className="absolute right-[1.6cqw] bottom-[1.4cqw]" style={{ width: "4cqw", height: "4cqw" }} strokeWidth={1.6} />
              </M.div>
            );
          })}
        </M.div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const gap = 0.35;
    const w = (11.7 - gap * (args.items.length - 1)) / args.items.length;
    args.items.forEach((it, i) => {
      const x = 0.8 + i * (w + gap);
      slide.addShape("rect", { x, y: 2.2, w, h: 4.7, fill: { color: hex(ctx.theme.secondary ?? "#8b5a2b"), transparency: 86 }, line: { type: "none" } as never });
      slide.addText([{ text: it.title, options: { bold: true, fontSize: 16, align: "center", breakLine: true } }, { text: plain(it.text), options: { fontSize: 12 } }], { x: x + 0.25, y: 2.4, w: w - 0.5, h: 4.3, valign: "top", fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
      // The corner icon, drawn as a picture before the export.
      const icon = pptxImage(`icon:${it.icon?.trim() || "sparkles"}`, { x: x + w - 0.85, y: 6.1, w: 0.6, h: 0.6 });
      if (icon) slide.addImage(icon);
    });
  },
});
