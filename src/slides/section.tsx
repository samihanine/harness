import { z } from "zod";
import { H, M, W, defineTemplate, ease, hex, pad, plain, reveal, rise, stagger } from "./kit";

/**
 * Transition between parts: the outline with the current part highlighted (its number large), and an
 * optional picture sliding in on the right. Starts a new part.
 */
export const section = defineTemplate({
  id: "section",
  description: "Transition slide starting a new part of the deck (shows the outline, current part highlighted). Optional picture URL.",
  section: true,
  bare: true,
  args: z.object({ title: z.string(), subtitle: z.string().optional(), image: z.string().optional() }),
  Component: ({ args, deck }) => (
    <div className={`grid size-full ${args.image ? "grid-cols-[50fr_50fr]" : "grid-cols-1"}`}>
      <M.div variants={stagger(0.07, 0.1)} className="flex flex-col justify-center gap-[0.9cqw] px-[7cqw]">
        <M.span variants={rise} className="text-[9cqw] leading-none font-bold tracking-tighter" style={{ color: "var(--s-primary)", fontFamily: "var(--s-title)" }}>
          {pad(deck.section + 1)}
        </M.span>
        {deck.sections.map((s, i) => (
          <M.div
            key={i}
            variants={{ hidden: { opacity: 0, x: "-2cqw" }, show: { opacity: i === deck.section ? 1 : 0.28, x: 0, transition: { duration: 0.7, ease } } }}
            className={i === deck.section ? "text-[3.4cqw] leading-tight font-bold" : "text-[1.5cqw] font-medium"}
            style={{ fontFamily: "var(--s-title)" }}
          >
            {i === deck.section ? s : `${pad(i + 1)}  ${s}`}
          </M.div>
        ))}
        {args.subtitle && <M.p variants={rise} className="mt-[1cqw] max-w-[38cqw] text-[1.5cqw] opacity-70">{args.subtitle}</M.p>}
      </M.div>
      {args.image && (
        <M.div variants={reveal} className="overflow-hidden">
          <M.img src={args.image} alt="" className="size-full object-cover" variants={{ hidden: { scale: 1.12 }, show: { scale: 1, transition: { duration: 2.4, ease } } }} />
        </M.div>
      )}
    </div>
  ),
  pptx: (slide, args, { theme, deck }) => {
    const w = args.image ? W / 2 : W;
    if (args.image) slide.addImage({ path: args.image, x: w, y: 0, w: W - w, h: H, sizing: { type: "cover", w: W - w, h: H } });
    slide.addText(pad(deck.section + 1), { x: 0.9, y: 1.4, w: 4, h: 1.5, fontSize: 88, bold: true, color: hex(theme.primary), fontFace: theme.fontTitle });
    slide.addText(plain(args.title), { x: 0.9, y: 3.0, w: w - 1.4, h: 1.2, fontSize: 36, bold: true, color: hex(theme.text), fontFace: theme.fontTitle });
    if (args.subtitle) slide.addText(args.subtitle, { x: 0.9, y: 4.3, w: w - 1.4, h: 0.8, fontSize: 16, color: hex(theme.text), fontFace: theme.fontText, transparency: 30 });
  },
});
