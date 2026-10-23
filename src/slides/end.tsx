import { z } from "zod";
import { M, Rich, Rule, W, defineTemplate, hex, plain, rise, stagger } from "./kit";

/** Last slide: thanks / next steps, centered, on the accent color. */
export const end = defineTemplate({
  id: "end",
  description: "Closing slide on the accent color: a title (e.g. \"Thank you\" / \"Questions?\") and an optional line (contact, next step).",
  kind: "end",
  bare: true,
  args: z.object({ title: z.string(), text: z.string().optional() }),
  Component: ({ args }) => (
    <M.div variants={stagger(0.14, 0.2)} className="flex size-full flex-col items-center justify-center gap-[1.8cqw] text-center" style={{ background: "var(--s-primary)", color: "#fff" }}>
      <M.h2 variants={rise} className="text-[6cqw] leading-none font-bold tracking-tight" style={{ fontFamily: "var(--s-title)" }}>
        {args.title}
      </M.h2>
      <div className="[&>div]:bg-white!">
        <Rule className="mx-auto" />
      </div>
      {args.text && <M.p variants={rise} className="max-w-[50cqw] text-[1.7cqw] opacity-85"><Rich text={args.text} /></M.p>}
    </M.div>
  ),
  pptx: (slide, args, { theme }) => {
    slide.background = { color: hex(theme.primary) };
    slide.addText(args.title, { x: 0, y: 2.4, w: W, h: 1.5, fontSize: 54, bold: true, color: "FFFFFF", align: "center", fontFace: theme.fontTitle });
    if (args.text) slide.addText(plain(args.text), { x: 1.5, y: 4.2, w: W - 3, h: 0.8, fontSize: 18, color: "FFFFFF", align: "center", fontFace: theme.fontText });
  },
});
