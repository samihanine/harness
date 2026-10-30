import { useState } from "react";
import { z } from "zod";
import { Frame, M, Rich, STAGE, deep, defineTemplate, ease, faint, hex, mix, plain, pptxPath, pptxTitle } from "./kit";

const { w: SW, h: SH } = STAGE;
/** A level's fill: the top the darkest (the accent darkened), the middle the accent, the base the palest. */
const shade = (i: number, n: number) => {
  const t = i / Math.max(1, n - 1);
  return t < 0.5 ? `color-mix(in srgb, ${deep} ${Math.round(80 - t * 160)}%, var(--s-primary))` : `color-mix(in srgb, var(--s-primary) ${Math.round(100 - (t - 0.5) * 120)}%, var(--s-bg))`;
};
/** A level's text color: white on the dark shades. */
const dark = (i: number, n: number) => i / Math.max(1, n - 1) < 0.5;
const ink = (i: number, n: number) => (dark(i, n) ? "white" : "var(--s-text)");

/** The pyramid's shape in the content box (STAGE units): its slices, the outward side of its left edge. */
function geometry(n: number) {
  const cx = 42;
  const top = 0.5;
  const base = SH - 0.5;
  const B = 30;
  const half = (y: number) => (B * (y - top)) / (base - top);
  const hh = (base - top) / n;
  const gap = 0.3;
  const y0 = (i: number) => top + i * hh + (i ? gap / 2 : 0);
  const y1 = (i: number) => top + (i + 1) * hh - (i < n - 1 ? gap / 2 : 0);
  // The left edge's outward normal and its angle (brackets along it).
  const len = Math.hypot(B, base - top);
  const normal = { x: -(base - top) / len, y: -B / len };
  const angle = (Math.atan2(-(base - top), B) * 180) / Math.PI;
  const along = (y: number, d: number) => ({ x: cx - half(y) + normal.x * d, y: y + normal.y * d });
  /** Where a level's text goes: the top slice is a triangle, its text sits low where it is wide enough. */
  const label = (i: number) => {
    const y = i ? (y0(i) + y1(i)) / 2 + hh * 0.08 : y1(i) - 2.2;
    return { y, width: 2 * half(y) * 0.9 };
  };
  /** A bracket over levels from..to (1 = the top): its corners and the middle of its label. */
  const bracket = (from: number, to: number, k: number) => {
    const a = y0(Math.max(0, Math.min(n - 1, from - 1))) + 0.4;
    const b = y1(Math.max(0, Math.min(n - 1, to - 1))) - 0.4;
    return { points: [along(a, 1.3), along(a, 2.2), along(b, 2.2), along(b, 1.3)], label: along((a + b) / 2, 3.4 + (k % 2) * 1.8) };
  };
  return { cx, half, hh, y0, y1, angle, label, bracket };
}

/**
 * A pyramid of levels (e.g. users by maturity): slices of a triangle from the top, labeled brackets
 * along its left edge (levels from..to), notes on the right; the levels build up from the base
 * unless `animated` is false. A level under the mouse lights up.
 */
export const pyramid = defineTemplate({
  id: "pyramid",
  description: "A pyramid of 2-6 levels from the top (the narrowest) to the base, each a title and an optional short text; brackets on the left (a label over levels from..to, 1 = the top, e.g. the tool they use); notes on the right (a note on a level). animated: false to draw it at once.",
  args: z.object({
    title: z.string(),
    levels: z.array(z.object({ title: z.string(), text: z.string().optional() })).min(2).max(6).describe("From the top (the narrowest) to the base"),
    brackets: z.array(z.object({ label: z.string(), from: z.number(), to: z.number() })).max(4).default([]).describe("On the left: a label over levels from..to (1 = the top), e.g. the tool they use"),
    notes: z.array(z.object({ level: z.number(), text: z.string() })).max(4).default([]).describe("On the right: a note on a level (1 = the top)"),
    animated: z.boolean().default(true).describe("false: the pyramid is drawn at once, without building up"),
  }),
  example: {
  title: "Niveaux d'usage",
  levels: [
    { title: "Data Analyst", text: "4e niveau" },
    { title: "Report Publisher", text: "3e niveau : publication contrôlée des rapports de la communauté" },
    { title: "Business Analyst", text: "2e niveau : Full Self Service et Analyze in Excel" },
    { title: "Insight Consumer", text: "1er niveau : rapports certifiés et personnalisation" },
  ],
  brackets: [
    { label: "Dataiku", from: 1, to: 1 },
    { label: "Power BI Desktop", from: 2, to: 2 },
    { label: "Power BI Online", from: 3, to: 4 },
  ],
  notes: [{ level: 3, text: "Les **key users** sont choisis parmi ce niveau" }],
},
  Component: ({ args }) => {
    const [hover, setHover] = useState<number>();
    const n = args.levels.length;
    const { cx, half, y0, y1, angle, label, bracket } = geometry(n);
    // Not animated: everything starts where it ends.
    const from = <T,>(v: T) => (args.animated ? v : false);
    return (
      <Frame title={args.title}>
        <div className="flex h-full items-center">
          <div className="relative w-full" style={{ height: `${SH}cqw` }}>
            <svg viewBox={`0 0 ${SW} ${SH}`} className="absolute inset-0 size-full overflow-visible" fill="none">
              {args.levels.map((_, i) => {
                const a = y0(i);
                const b = y1(i);
                return (
                  <M.polygon
                    key={i}
                    points={`${cx - half(a)},${a} ${cx + half(a)},${a} ${cx + half(b)},${b} ${cx - half(b)},${b}`}
                    style={{ fill: shade(i, n), filter: hover === i ? "brightness(1.08)" : undefined }}
                    initial={from({ opacity: 0, y: 2 })}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: 0.2 + (n - 1 - i) * 0.22, duration: 0.6, ease }}
                    onMouseEnter={() => setHover(i)}
                    onMouseLeave={() => setHover(undefined)}
                  />
                );
              })}
              {args.brackets.map((br, k) => {
                const { points, label: m } = bracket(br.from, br.to, k);
                return (
                  <g key={k}>
                    <M.path d={points.map((p, j) => `${j ? "L" : "M"}${p.x} ${p.y}`).join(" ")} style={{ stroke: deep }} strokeWidth={0.18} initial={from({ pathLength: 0 })} animate={{ pathLength: 1 }} transition={{ delay: 1.2 + k * 0.2, duration: 0.7 }} />
                    <M.text x={m.x} y={m.y} transform={`rotate(${angle} ${m.x} ${m.y})`} textAnchor="middle" dominantBaseline="middle" style={{ fontSize: 1.05, fontWeight: 700, letterSpacing: 0.12, fill: deep, fontFamily: "var(--s-title)" }} initial={from({ opacity: 0 })} animate={{ opacity: 1 }} transition={{ delay: 1.5 + k * 0.2 }}>
                      {br.label.toUpperCase()}
                    </M.text>
                  </g>
                );
              })}
              {args.notes.map((note, k) => {
                const i = Math.min(n - 1, Math.max(0, note.level - 1));
                const y = (y0(i) + y1(i)) / 2;
                return <M.path key={k} d={`M${cx + half(y) + 0.8} ${y} H${SW - 17}`} style={{ stroke: faint(35) }} strokeWidth={0.1} initial={from({ pathLength: 0 })} animate={{ pathLength: 1 }} transition={{ delay: 1.6 + k * 0.2, duration: 0.6 }} />;
              })}
            </svg>
            {args.levels.map((l, i) => {
              const { y, width } = label(i);
              return (
                <M.div key={i} initial={from({ opacity: 0 })} animate={{ opacity: 1 }} transition={{ delay: 0.5 + (n - 1 - i) * 0.22, duration: 0.6 }} className="pointer-events-none absolute flex -translate-x-1/2 -translate-y-1/2 flex-col items-center text-center leading-tight" style={{ left: `${cx}cqw`, top: `${y}cqw`, width: `${width}cqw`, color: ink(i, n) }}>
                  <span className="text-[1.3cqw] font-bold tracking-wide uppercase">{l.title}</span>
                  {l.text && i > 0 && <span className="mt-[0.2cqw] text-[1cqw] opacity-85"><Rich text={l.text} /></span>}
                </M.div>
              );
            })}
            {args.notes.map((note, k) => {
              const i = Math.min(n - 1, Math.max(0, note.level - 1));
              return (
                <M.div key={k} initial={from({ opacity: 0, x: "1cqw" })} animate={{ opacity: 1, x: 0 }} transition={{ delay: 1.9 + k * 0.2, duration: 0.5 }} className="absolute right-0 -translate-y-1/2 text-[1.35cqw] leading-snug" style={{ top: `${(y0(i) + y1(i)) / 2}cqw`, width: "16cqw" }}>
                  <Rich text={note.text} />
                </M.div>
              );
            })}
          </div>
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const { theme } = ctx;
    const n = args.levels.length;
    const { cx, half, y0, y1, angle, label, bracket } = geometry(n);
    // The content box on the slide: 1 unit is 1 cqw, 1/100 of the slide's width.
    const k = 13.333 / 100;
    const at = (p: { x: number; y: number }) => ({ x: 0.8 + p.x * k, y: 2.15 + p.y * k });
    const strong = mix(theme.primary, theme.text, 0.62);
    const fill = (i: number) => {
      const t = i / Math.max(1, n - 1);
      return t < 0.5 ? mix(strong, theme.primary, (80 - t * 160) / 100) : mix(theme.primary, theme.background, (100 - (t - 0.5) * 120) / 100);
    };
    args.levels.forEach((l, i) => {
      const a = y0(i);
      const b = y1(i);
      pptxPath(slide, [{ x: cx - half(a), y: a }, { x: cx + half(a), y: a }, { x: cx + half(b), y: b }, { x: cx - half(b), y: b }].map(at), { color: fill(i), width: 0.1, fill: fill(i) });
      const { y, width } = label(i);
      const c = at({ x: cx, y });
      slide.addText(
        [{ text: l.title.toUpperCase(), options: { bold: true, fontSize: 12, breakLine: !!(l.text && i) } }, ...(l.text && i ? [{ text: plain(l.text), options: { fontSize: 9 } }] : [])],
        { x: c.x - (width * k) / 2, y: c.y - 0.35, w: width * k, h: 0.7, align: "center", valign: "middle", color: dark(i, n) ? "FFFFFF" : hex(theme.text), fontFace: theme.fontText },
      );
    });
    args.brackets.forEach((br, j) => {
      const { points, label: m } = bracket(br.from, br.to, j);
      pptxPath(slide, points.map(at), { color: strong, width: 1.25 });
      const c = at(m);
      slide.addText(br.label.toUpperCase(), { x: c.x - 1.6, y: c.y - 0.15, w: 3.2, h: 0.3, rotate: 360 + angle, wrap: false, align: "center", valign: "middle", fontSize: 9, bold: true, charSpacing: 1, color: strong, fontFace: theme.fontTitle });
    });
    args.notes.forEach((note) => {
      const i = Math.min(n - 1, Math.max(0, note.level - 1));
      const y = (y0(i) + y1(i)) / 2;
      pptxPath(slide, [at({ x: cx + half(y) + 0.8, y }), at({ x: SW - 17, y })], { color: mix(theme.text, theme.background, 0.35), width: 0.75 });
      const p = at({ x: SW - 16, y });
      slide.addText(plain(note.text), { x: p.x, y: p.y - 0.4, w: 16 * k, h: 0.8, fontSize: 12, valign: "middle", color: hex(theme.text), fontFace: theme.fontText });
    });
  },
});
