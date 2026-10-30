import { z } from "zod";
import { Frame, M, Rich, defineTemplate, hex, pad, plain, pptxTitle, rise, runs, stagger } from "./kit";

/** Numbered points (title + text), with an optional lead sentence. */
export const list = defineTemplate({
  id: "list",
  description: "Numbered points: title, optional lead sentence, 2-6 items (title + short text).",
  args: z.object({ title: z.string(), lead: z.string().optional(), items: z.array(z.object({ title: z.string(), text: z.string().optional() })).min(1).max(8) }),
  Component: ({ args }) => (
    <Frame title={args.title}>
      <div className="grid h-full grid-cols-[34fr_66fr] gap-[4cqw]">
        <M.p variants={rise} className="text-[1.9cqw] leading-snug font-medium" style={{ fontFamily: "var(--s-title)" }}>
          {args.lead && <Rich text={args.lead} />}
        </M.p>
        <M.ol variants={stagger(0.1, 0.35)} className="flex flex-col justify-center gap-[1.6cqw]">
          {args.items.map((item, i) => (
            <M.li key={i} variants={rise} className="grid grid-cols-[4cqw_1fr] items-baseline border-b pb-[1.2cqw]" style={{ borderColor: "color-mix(in srgb, var(--s-text) 12%, transparent)" }}>
              <span className="text-[1.3cqw] font-bold tabular-nums" style={{ color: "var(--s-primary)" }}>{pad(i + 1)}</span>
              <span className="flex flex-col gap-[0.3cqw]">
                <span className="text-[1.6cqw] font-semibold"><Rich text={item.title} /></span>
                {item.text && <span className="text-[1.25cqw] leading-snug opacity-70"><Rich text={item.text} /></span>}
              </span>
            </M.li>
          ))}
        </M.ol>
      </div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    if (args.lead) slide.addText(runs(args.lead, ctx.theme), { x: 0.8, y: 2.3, w: 3.8, h: 3, fontSize: 18, fontFace: ctx.theme.fontTitle, color: hex(ctx.theme.text), valign: "top" });
    const h = Math.min(1.1, 4.6 / args.items.length);
    args.items.forEach((item, i) => {
      const y = 2.3 + i * h;
      slide.addText(pad(i + 1), { x: 5.2, y, w: 0.6, h: 0.4, fontSize: 13, bold: true, color: hex(ctx.theme.primary), fontFace: ctx.theme.fontText });
      slide.addText([{ text: plain(item.title), options: { bold: true, fontSize: 16, breakLine: true } }, ...(item.text ? [{ text: plain(item.text), options: { fontSize: 12, transparency: 30 } }] : [])], { x: 5.9, y, w: 6.6, h: h - 0.1, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text), valign: "top" });
    });
  },
});
