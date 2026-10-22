import { z } from "zod";
import { Frame, M, Rich, defineTemplate, hex, plain, pptxTitle, rise, stagger } from "./kit";

/** Cards side by side (2-4): a title and a text each, an optional lead above. */
export const cards = defineTemplate({
  id: "cards",
  description: "2-4 cards side by side (title + text each), optional lead sentence.",
  args: z.object({ title: z.string(), eyebrow: z.string().optional(), lead: z.string().optional(), items: z.array(z.object({ title: z.string(), text: z.string() })).min(1).max(4) }),
  Component: ({ args }) => (
    <Frame eyebrow={args.eyebrow} title={args.title}>
      <div className="flex h-full flex-col gap-[2.4cqw]">
        {args.lead && <M.p variants={rise} className="max-w-[60cqw] text-[1.6cqw] leading-snug opacity-80"><Rich text={args.lead} /></M.p>}
        <M.div variants={stagger(0.12, 0.3)} className="grid flex-1 gap-[1.6cqw]" style={{ gridTemplateColumns: `repeat(${args.items.length}, 1fr)` }}>
          {args.items.map((c, i) => (
            <M.div key={i} variants={rise} className="flex flex-col gap-[1cqw] rounded-[1.2cqw] p-[2cqw]" style={{ background: "color-mix(in srgb, var(--s-primary) 7%, var(--s-bg))", borderTop: "0.35cqw solid var(--s-primary)" }}>
              <span className="text-[1.7cqw] font-bold" style={{ fontFamily: "var(--s-title)" }}><Rich text={c.title} /></span>
              <span className="text-[1.25cqw] leading-snug opacity-75"><Rich text={c.text} /></span>
            </M.div>
          ))}
        </M.div>
      </div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx, args.eyebrow);
    const top = args.lead ? 3.0 : 2.3;
    if (args.lead) slide.addText(plain(args.lead), { x: 0.8, y: 2.2, w: 11.5, h: 0.6, fontSize: 16, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    const gap = 0.25;
    const w = (11.7 - gap * (args.items.length - 1)) / args.items.length;
    args.items.forEach((c, i) => {
      const x = 0.8 + i * (w + gap);
      slide.addShape("roundRect", { x, y: top, w, h: 6.8 - top, fill: { color: hex(ctx.theme.primary), transparency: 92 }, line: { color: hex(ctx.theme.primary), transparency: 100 }, rectRadius: 0.12 });
      slide.addShape("rect", { x, y: top, w, h: 0.06, fill: { color: hex(ctx.theme.primary) }, line: { color: hex(ctx.theme.primary) } });
      slide.addText([{ text: plain(c.title), options: { bold: true, fontSize: 18, breakLine: true } }, { text: plain(c.text), options: { fontSize: 13 } }], { x: x + 0.2, y: top + 0.25, w: w - 0.4, h: 6.4 - top, valign: "top", fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
    });
  },
});
