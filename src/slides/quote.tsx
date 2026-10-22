import { z } from "zod";
import { H, M, Rich, W, defineTemplate, hex, plain, rise, stagger } from "./kit";

/** One sentence to remember, large, with its author. */
export const quote = defineTemplate({
  id: "quote",
  description: "One key sentence or quote, large, with an optional author.",
  bare: true,
  args: z.object({ text: z.string(), author: z.string().optional() }),
  Component: ({ args }) => (
    <M.div variants={stagger(0.15, 0.2)} className="flex size-full flex-col justify-center gap-[2cqw] px-[12cqw]">
      <M.span variants={rise} className="text-[10cqw] leading-[0.6] font-bold" style={{ color: "var(--s-primary)", fontFamily: "var(--s-title)" }}>“</M.span>
      <M.p variants={rise} className="text-[3.4cqw] leading-[1.18] font-semibold tracking-tight" style={{ fontFamily: "var(--s-title)" }}>
        <Rich text={args.text} />
      </M.p>
      {args.author && <M.p variants={rise} className="text-[1.4cqw] font-medium opacity-60">— {args.author}</M.p>}
    </M.div>
  ),
  pptx: (slide, args, { theme }) => {
    slide.addText("“", { x: 1.5, y: 1.2, w: 2, h: 1.4, fontSize: 110, bold: true, color: hex(theme.primary), fontFace: theme.fontTitle });
    slide.addText(plain(args.text), { x: 1.5, y: 2.5, w: W - 3, h: 2.6, fontSize: 32, bold: true, color: hex(theme.text), fontFace: theme.fontTitle, valign: "top" });
    if (args.author) slide.addText(`— ${args.author}`, { x: 1.5, y: H - 1.9, w: W - 3, h: 0.5, fontSize: 15, color: hex(theme.text), transparency: 40, fontFace: theme.fontText });
  },
});
