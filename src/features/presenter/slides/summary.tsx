import { z } from "zod";
import { M, defineTemplate, ease, hex, pad, plain, rise, stagger } from "./kit";

/**
 * Summary page: the deck's outline, made from its part and subpart pages (nothing to keep in sync),
 * numbered; `current` lights the part about to start (a summary before each part).
 */
export const summary = defineTemplate({
  id: "summary",
  description: "Summary page: the outline of the deck, filled automatically from its section and subtitle pages. current: number of the part to highlight (a summary repeated before each part).",
  kind: "summary",
  bare: true,
  args: z.object({ title: z.string().default("Summary"), current: z.number().optional().describe("Part to highlight, from 1") }),
  Component: ({ args, deck }) => {
    const parts = deck.outline.filter((p) => p.title);
    const two = parts.length > 4;
    return (
      <M.div variants={stagger(0.08, 0.2)} className="grid size-full grid-cols-[34fr_66fr] gap-[5cqw] px-[7cqw] py-[6cqw]">
        <div className="flex flex-col justify-center gap-[1.2cqw]">
          <M.h2 variants={rise} className="text-[4.4cqw] leading-none font-bold tracking-tight" style={{ fontFamily: "var(--s-title)" }}>
            {args.title}
          </M.h2>
          <M.div variants={{ hidden: { scaleX: 0 }, show: { scaleX: 1, transition: { duration: 0.9, ease } } }} className="h-[0.22cqw] w-[6cqw] origin-left rounded-full" style={{ background: "var(--s-primary)" }} />
        </div>
        <M.ol variants={stagger(0.1, 0.4)} className={`grid content-center gap-x-[3cqw] gap-y-[1.8cqw] ${two ? "grid-cols-2" : "grid-cols-1"}`}>
          {parts.map((p, i) => {
            const lit = args.current === undefined || args.current === i + 1;
            return (
              <M.li key={i} variants={rise} className="flex gap-[1.4cqw] border-b pb-[1.2cqw]" style={{ borderColor: "color-mix(in srgb, var(--s-text) 12%, transparent)", opacity: lit ? 1 : 0.35 }}>
                <span className="text-[1.6cqw] font-bold tabular-nums" style={{ color: "var(--s-primary)", fontFamily: "var(--s-title)" }}>
                  {pad(i + 1)}
                </span>
                <span className="flex flex-col gap-[0.4cqw]">
                  <span className="text-[1.9cqw] leading-tight font-semibold" style={{ fontFamily: "var(--s-title)" }}>
                    {p.title}
                  </span>
                  {p.subsections.length > 0 && <span className="text-[1.05cqw] leading-snug opacity-60">{p.subsections.join(" · ")}</span>}
                </span>
              </M.li>
            );
          })}
          {!parts.length && <M.li variants={rise} className="text-[1.4cqw] opacity-50">Add section pages: the summary lists them.</M.li>}
        </M.ol>
      </M.div>
    );
  },
  pptx: (slide, args, { theme, deck }) => {
    const parts = deck.outline.filter((p) => p.title);
    slide.addText(args.title, { x: 0.9, y: 2.6, w: 3.8, h: 1.2, fontSize: 40, bold: true, fontFace: theme.fontTitle, color: hex(theme.text) });
    slide.addShape("rect", { x: 0.9, y: 3.9, w: 0.8, h: 0.05, fill: { color: hex(theme.primary) }, line: { color: hex(theme.primary) } });
    const h = Math.min(1, 5.4 / Math.max(1, parts.length));
    parts.forEach((p, i) => {
      const y = 1.2 + i * h;
      const lit = args.current === undefined || args.current === i + 1;
      slide.addText(pad(i + 1), { x: 5.2, y, w: 0.7, h: 0.5, fontSize: 18, bold: true, color: hex(theme.primary), fontFace: theme.fontTitle, transparency: lit ? 0 : 60 });
      slide.addText([{ text: plain(p.title), options: { bold: true, fontSize: 18, breakLine: true } }, ...(p.subsections.length ? [{ text: p.subsections.join(" · "), options: { fontSize: 11 } }] : [])], { x: 5.9, y, w: 6.5, h: h - 0.1, valign: "top", fontFace: theme.fontText, color: hex(theme.text), transparency: lit ? 0 : 60 });
    });
  },
});
