import { useState } from "react";
import { z } from "zod";
import { Frame, M, Rich, deep, defineTemplate, ease, faint, hex, mix, plain, pptxPath, pptxTitle, rise, tint } from "./kit";

/**
 * Causes and their effects, row by row: a numbered cause, a curved bridge carrying a dot, its effect
 * (the issues page of older versions: issue → solution). A row under the mouse lights up.
 */
export const causality = defineTemplate({
  id: "causality",
  description: "Cause → effect pairs (2-6), row by row joined by a curved arrow (e.g. a problem and its consequence, a risk and how it is handled). Optional column headings (default Cause / Effet).",
  args: z.object({
    title: z.string(),
    headings: z.object({ left: z.string(), right: z.string() }).default({ left: "Cause", right: "Effet" }),
    items: z.array(z.object({ cause: z.string(), effect: z.string() })).min(2).max(6),
  }),
  example: {
  title: "Causes et effets",
  headings: { left: "Cause", right: "Effet" },
  items: [
    { cause: "Des rapports **peu documentés**", effect: "Les mêmes questions reviennent aux key users" },
    { cause: "Des modèles très **volumineux**", effect: "Des requêtes lentes et des refresh qui échouent" },
    { cause: "Une licence **Pro** seulement", effect: "Pas d'export API : tout passe par le navigateur" },
    { cause: "Des définitions **différentes** par équipe", effect: "Deux chiffres pour un même indicateur" },
  ],
},
  Component: ({ args }) => {
    const [hover, setHover] = useState<number>();
    const dense = args.items.length > 4;
    return (
      <Frame title={args.title}>
        <div className="grid h-full grid-cols-[3cqw_1fr_9cqw_1fr] content-center items-center gap-x-[1cqw] gap-y-[1.3cqw]">
          <span />
          <M.span variants={rise} className="text-[1cqw] font-semibold tracking-[0.16em] uppercase opacity-55">{args.headings.left}</M.span>
          <span />
          <M.span variants={rise} className="text-[1cqw] font-semibold tracking-[0.16em] uppercase" style={{ color: deep }}>{args.headings.right}</M.span>
          {args.items.map((it, i) => {
            const delay = 0.3 + i * 0.3;
            const on = hover === i;
            return [
              <M.span key={`n${i}`} initial={{ opacity: 0, scale: 0.5 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay, duration: 0.4 }} className="grid size-[2.4cqw] place-items-center rounded-full text-[1.1cqw] font-bold" style={{ background: on ? "var(--s-primary)" : tint(25), color: on ? "white" : deep }}>
                {i + 1}
              </M.span>,
              <M.div key={`c${i}`} onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(undefined)} initial={{ opacity: 0, x: "-1.5cqw" }} animate={{ opacity: 1, x: 0 }} transition={{ duration: 0.6, delay, ease }} className={`rounded-[0.9cqw] px-[1.4cqw] leading-snug transition-colors ${dense ? "py-[0.9cqw] text-[1.4cqw]" : "py-[1.4cqw] text-[1.65cqw]"}`} style={{ background: on ? faint(9) : faint(5) }}>
                <Rich text={it.cause} />
              </M.div>,
              <svg key={`b${i}`} viewBox="0 0 90 30" className="h-[3cqw] w-full overflow-visible" fill="none" preserveAspectRatio="none">
                <M.path d="M2 15 C30 -4 60 34 86 15" style={{ stroke: on ? deep : "var(--s-primary)" }} strokeWidth={on ? 3.5 : 2.5} strokeLinecap="round" initial={{ pathLength: 0 }} animate={{ pathLength: 1 }} transition={{ delay: delay + 0.35, duration: 0.6, ease }} />
                <M.path d="M80 9 L88 15 L80 21" style={{ stroke: "var(--s-primary)" }} strokeWidth={2.5} strokeLinecap="round" strokeLinejoin="round" initial={{ opacity: 0 }} animate={{ opacity: 1 }} transition={{ delay: delay + 0.9 }} />
                <circle r={3.2} style={{ fill: deep }} opacity={0}>
                  <animate attributeName="opacity" to="1" begin={`${delay + 1}s`} dur="0.2s" fill="freeze" />
                  <animateMotion dur="1.8s" begin={`${delay + 1}s`} repeatCount="indefinite" path="M2 15 C30 -4 60 34 86 15" />
                </circle>
              </svg>,
              <M.div key={`e${i}`} onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(undefined)} initial={{ opacity: 0, x: "1.5cqw" }} animate={{ opacity: 1, x: 0 }} transition={{ duration: 0.6, delay: delay + 0.8, ease }} className={`rounded-[0.9cqw] border-l-[0.35cqw] px-[1.4cqw] leading-snug font-medium transition-shadow ${dense ? "py-[0.9cqw] text-[1.4cqw]" : "py-[1.4cqw] text-[1.65cqw]"}`} style={{ borderColor: "var(--s-primary)", background: tint(on ? 26 : 14), boxShadow: on ? "0 0.3cqw 1cqw rgba(0,0,0,0.1)" : undefined }}>
                <Rich text={it.effect} />
              </M.div>,
            ];
          })}
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const { theme } = ctx;
    const text = hex(theme.text);
    const strong = mix(theme.primary, theme.text, 0.62);
    slide.addText(args.headings.left.toUpperCase(), { x: 1.35, y: 2.0, w: 4.6, h: 0.35, fontSize: 10, bold: true, color: text, transparency: 45, fontFace: theme.fontText });
    slide.addText(args.headings.right.toUpperCase(), { x: 7.5, y: 2.0, w: 5.0, h: 0.35, fontSize: 10, bold: true, color: strong, fontFace: theme.fontText });
    const h = Math.min(0.8, 4.2 / args.items.length - 0.12);
    args.items.forEach((it, i) => {
      const y = 2.45 + i * (h + 0.12);
      slide.addShape("ellipse", { x: 0.8, y: y + h / 2 - 0.17, w: 0.34, h: 0.34, fill: { color: mix(theme.primary, theme.background, 0.25) }, line: { type: "none" } as never });
      slide.addText(String(i + 1), { x: 0.8, y: y + h / 2 - 0.17, w: 0.34, h: 0.34, fontSize: 10, bold: true, align: "center", valign: "middle", color: strong });
      slide.addShape("roundRect", { x: 1.35, y, w: 4.6, h, fill: { color: mix(theme.text, theme.background, 0.05) }, line: { type: "none" } as never, rectRadius: 0.1 });
      slide.addText(plain(it.cause), { x: 1.5, y, w: 4.35, h, fontSize: 13, valign: "middle", color: text, fontFace: theme.fontText });
      const m = y + h / 2;
      pptxPath(slide, [{ x: 6.1, y: m }, { x: 7.3, y: m, c1: { x: 6.5, y: m - 0.25 }, c2: { x: 6.9, y: m + 0.25 } }], { color: hex(theme.primary), width: 2, arrow: true });
      slide.addShape("rect", { x: 7.5, y, w: 5.0, h, fill: { color: mix(theme.primary, theme.background, 0.14) }, line: { type: "none" } as never });
      slide.addShape("rect", { x: 7.5, y, w: 0.05, h, fill: { color: hex(theme.primary) }, line: { type: "none" } as never });
      slide.addText(plain(it.effect), { x: 7.7, y, w: 4.7, h, fontSize: 13, valign: "middle", color: text, fontFace: theme.fontText });
    });
  },
});
