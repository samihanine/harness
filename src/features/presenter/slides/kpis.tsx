import { z } from "zod";
import { Frame, M, defineTemplate, hex, pptxTitle, rise, stagger } from "./kit";

/** Key figures: big values with a label and an optional trend note. */
export const kpis = defineTemplate({
  id: "kpis",
  description: "2-4 key figures: value (text, e.g. \"12.4 M€\"), label, optional note (e.g. \"+8 % vs 2024\").",
  args: z.object({ title: z.string(), items: z.array(z.object({ value: z.string(), label: z.string(), note: z.string().optional() })).min(1).max(4) }),
  Component: ({ args }) => (
    <Frame title={args.title}>
      <M.div variants={stagger(0.12, 0.3)} className="grid h-full content-center items-start gap-[2cqw]" style={{ gridTemplateColumns: `repeat(${args.items.length}, 1fr)` }}>
        {args.items.map((k, i) => (
          <M.div key={i} variants={rise} className="flex flex-col gap-[0.6cqw] border-l-[0.3cqw] pl-[2cqw]" style={{ borderColor: "var(--s-primary)" }}>
            <span className="text-[5.2cqw] leading-none font-bold tracking-tight tabular-nums" style={{ fontFamily: "var(--s-title)" }}>{k.value}</span>
            <span className="text-[1.5cqw] font-medium">{k.label}</span>
            {k.note && <span className="text-[1.15cqw] font-semibold" style={{ color: "var(--s-primary)" }}>{k.note}</span>}
          </M.div>
        ))}
      </M.div>
    </Frame>
  ),
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const w = 11.7 / args.items.length;
    args.items.forEach((k, i) => {
      const x = 0.8 + i * w;
      slide.addShape("rect", { x, y: 3.0, w: 0.05, h: 2.2, fill: { color: hex(ctx.theme.primary) }, line: { color: hex(ctx.theme.primary) } });
      slide.addText(k.value, { x: x + 0.25, y: 2.9, w: w - 0.4, h: 1.2, fontSize: 48, bold: true, fontFace: ctx.theme.fontTitle, color: hex(ctx.theme.text) });
      slide.addText(k.label, { x: x + 0.25, y: 4.1, w: w - 0.4, h: 0.5, fontSize: 16, fontFace: ctx.theme.fontText, color: hex(ctx.theme.text) });
      if (k.note) slide.addText(k.note, { x: x + 0.25, y: 4.6, w: w - 0.4, h: 0.4, fontSize: 13, bold: true, fontFace: ctx.theme.fontText, color: hex(ctx.theme.primary) });
    });
  },
});
