/*
 * Toile de Jouy flowers (an engraving in the deck's accent color) for the title pages: around the
 * title (ToileFlowers) and one rose spray growing from the subpart trail (ToileSpray).
 * Pure SVG + CSS animations (.toile-* in styles.css): stems draw themselves, leaves unfold, roses
 * bloom, then everything breathes and sways very slightly; a few petals drift down. Still slides
 * (thumbnails, PDF) show the final drawing. Drawn in a 1600×900 box (the 16:9 slide); the right
 * branch is the left one mirrored. Adapted from the Hub presentation (toile.tsx).
 */

import type { CSSProperties } from "react";

type P = [number, number];
const n = (v: number) => Math.round(v * 10) / 10;
const deg = (rad: number) => (rad * 180) / Math.PI;

/** Cubic bezier: point and direction (radians) at t. */
const bezier = ([a, b, c, d]: P[], t: number) => {
  const u = 1 - t;
  const at = (i: 0 | 1) =>
    u * u * u * a[i] +
    3 * u * u * t * b[i] +
    3 * u * t * t * c[i] +
    t * t * t * d[i];
  const tan = (i: 0 | 1) =>
    3 * u * u * (b[i] - a[i]) +
    6 * u * t * (c[i] - b[i]) +
    3 * t * t * (d[i] - c[i]);
  return { p: [at(0), at(1)] as P, angle: Math.atan2(tan(1), tan(0)) };
};
const curve = ([a, b, c, d]: P[]) => `M${a} C${b} ${c} ${d}`;

type Anim = CSSProperties & { "--d"?: string; "--dur"?: string };
const delay = (d: number, dur?: number): Anim => ({
  "--d": `${d}s`,
  ...(dur ? { "--dur": `${dur}s` } : {}),
});
/** Idle motion phase: the later a motif appears (further along its branch), the later the gust reaches it. */
const gust = (d: number, dur: number) => delay(n(d * 0.7 - 30), dur);

/* ------------------------------------------------------------------ */
/* Engraved motifs (local coordinates, pointing up)                    */
/* ------------------------------------------------------------------ */

/** Loose petal (drifting petals). */
const PETAL =
  "M-20 -8C-40 -22 -37 -50 -14 -52C-6 -58 6 -58 14 -52C37 -50 40 -22 20 -8Z";

/** Parallel hatch strokes following a curve, offset step by step (engraving shade). */
const hatch = (count: number, f: (i: number) => string) =>
  Array.from({ length: count }, (_, i) => f(i));

/* Cabbage rose in three-quarter view, drawn back to front in a ~110 unit box. */
const ROSE_BACK = [
  "M-46 -4C-60 -30 -42 -58 -16 -50C-8 -64 18 -64 26 -50C50 -58 64 -30 48 -4Z",
  "M-46 -6C-66 4 -62 32 -38 40C-32 22 -30 8 -22 -2Z",
  "M46 -6C66 4 62 32 38 40C32 22 30 8 22 -2Z",
];
const ROSE_BACK_RIMS = [
  "M-16 -50C-12 -44 -4 -42 4 -44",
  "M26 -50C22 -45 18 -43 12 -43",
  "M-56 -24C-48 -20 -42 -16 -38 -10",
  "M56 -24C48 -20 42 -16 38 -10",
];
const ROSE_BACK_HATCH = [
  ...hatch(
    6,
    (i) =>
      `M${-30 + i * 5} -44C${-31 + i * 5} -38 ${-30 + i * 5} -32 ${-28 + i * 5} -27`,
  ),
  ...hatch(
    5,
    (i) =>
      `M${14 + i * 5} -46C${14 + i * 5} -40 ${15 + i * 5} -34 ${16 + i * 5} -29`,
  ),
  ...hatch(
    5,
    (i) =>
      `M${-52 + i * 3} ${4 + i * 5}C${-46 + i * 3} ${8 + i * 5} ${-40 + i * 3} ${10 + i * 5} ${-34 + i * 2} ${10 + i * 5}`,
  ),
];
const ROSE_CUP = "M-36 -10C-32 -38 32 -40 36 -10C24 -4 -24 -4 -36 -10Z";
const ROSE_HEART = [
  "M-26 -12C-24 -32 22 -34 26 -12",
  "M-16 -12C-14 -26 14 -28 18 -14C20 -4 6 2 -4 -2C-10 -5 -8 -14 0 -15C7 -16 9 -9 4 -7",
  "M-20 -14C-18 -22 -10 -26 -2 -27",
];
const ROSE_HEART_HATCH = hatch(
  7,
  (i) =>
    `M${-22 + i * 6} -14C${-21 + i * 6} -10 ${-20 + i * 6} -8 ${-19 + i * 6} -6`,
);
const ROSE_SIDES = [
  "M-36 -10C-46 6 -38 26 -18 32C-24 16 -24 4 -16 -6Z",
  "M36 -10C46 6 38 26 18 32C24 16 24 4 16 -6Z",
];
const ROSE_SIDES_HATCH = [
  ...hatch(
    4,
    (i) =>
      `M${-38 + i * 4} ${-2 + i * 2}C${-36 + i * 4} ${8 + i * 2} ${-32 + i * 3} ${16 + i * 2} ${-26 + i * 2} ${22 + i}`,
  ),
  ...hatch(
    3,
    (i) =>
      `M${38 - i * 4} ${-2 + i * 2}C${36 - i * 4} ${8 + i * 2} ${33 - i * 3} ${15 + i * 2} ${28 - i * 2} ${20 + i}`,
  ),
];
const ROSE_FRONT =
  "M-42 -8C-30 -2 -14 2 0 0C14 2 30 -2 42 -8C46 18 28 44 0 46C-28 44 -46 18 -42 -8Z";
const ROSE_FRONT_LIP = "M-42 -8C-34 -14 -16 -12 0 -6C16 -12 34 -14 42 -8";
const ROSE_FRONT_HATCH = [
  ...hatch(
    6,
    (i) =>
      `M${-38 + i * 2} ${2 + i * 3}C${-32 + i * 2} ${22 + i * 2} ${-16 + i} ${34 + i} ${-2} ${36 + i * 1.5}`,
  ),
  ...hatch(
    11,
    (i) =>
      `M${-26 + i * 5} ${-3 + Math.abs(i - 5) * 0.6}L${-25 + i * 5} ${5 - Math.abs(i - 5) * 0.4}`,
  ),
  ...hatch(
    3,
    (i) =>
      `M${34 - i * 3} ${6 + i * 4}C${30 - i * 3} ${20 + i * 3} ${22 - i * 2} ${30 + i * 2} ${14 - i} ${34 + i}`,
  ),
];
const ROSE_WRAP = "M-44 2C-40 30 -18 48 6 46C-12 38 -24 22 -26 2Z";
const ROSE_WRAP_HATCH = hatch(
  4,
  (i) =>
    `M${-40 + i * 3} ${12 + i * 4}C${-34 + i * 3} ${26 + i * 3} ${-24 + i * 2} ${36 + i * 2} ${-14 + i} ${40 + i}`,
);

/** Open cabbage rose (three-quarter view): back petals, cup with a swirled heart, rolled front petals. */
function Rose({
  x,
  y,
  s = 1,
  r = 0,
  d = 0,
}: {
  x: number;
  y: number;
  s?: number;
  r?: number;
  d?: number;
}) {
  return (
    <g
      transform={`translate(${x} ${y}) rotate(${r}) scale(${s})`}
     
    >
      <g className="toile-bloom" style={delay(d)}>
        <g className="toile-breathe" style={gust(d, 6 + (d % 2))}>
          {ROSE_BACK.map((p) => (
            <path key={p} d={p} className="toile-paper" />
          ))}
          {[...ROSE_BACK_RIMS, ...ROSE_BACK_HATCH].map((p, i) => (
            <path
              key={p}
              d={p}
              className={i < ROSE_BACK_RIMS.length ? undefined : "toile-hatch"}
            />
          ))}
          <path d={ROSE_CUP} className="toile-paper" />
          {ROSE_HEART.map((p) => (
            <path key={p} d={p} />
          ))}
          {ROSE_HEART_HATCH.map((p) => (
            <path key={p} d={p} className="toile-hatch" />
          ))}
          {ROSE_SIDES.map((p) => (
            <path key={p} d={p} className="toile-paper" />
          ))}
          {ROSE_SIDES_HATCH.map((p) => (
            <path key={p} d={p} className="toile-hatch" />
          ))}
          <path d={ROSE_FRONT} className="toile-paper" />
          <path d={ROSE_FRONT_LIP} />
          {ROSE_FRONT_HATCH.map((p) => (
            <path key={p} d={p} className="toile-hatch" />
          ))}
          <path d={ROSE_WRAP} className="toile-paper" />
          {ROSE_WRAP_HATCH.map((p) => (
            <path key={p} d={p} className="toile-hatch" />
          ))}
        </g>
      </g>
    </g>
  );
}

/** Closed bud with three sepals curling around it. */
function Bud({
  x,
  y,
  s = 1,
  r = 0,
  d = 0,
}: {
  x: number;
  y: number;
  s?: number;
  r?: number;
  d?: number;
}) {
  return (
    <g
      transform={`translate(${x} ${y}) rotate(${r}) scale(${s})`}
     
    >
      <g className="toile-grow" style={delay(d)}>
        <g className="toile-flutter" style={gust(d, 5.5)}>
          <path
            d="M0 0C-12 -8 -14 -26 0 -40C14 -26 12 -8 0 0Z"
            className="toile-paper"
          />
          <path d="M-9 -12C-6 -24 2 -32 8 -33" />
          <path d="M-6 -8C-6 -18 -2 -26 3 -30" className="toile-hatch" />
          <path d="M-3 -6C-3 -14 -1 -20 1 -24" className="toile-hatch" />
          <path
            d="M0 0C-8 -4 -15 -12 -17 -26C-11 -16 -5 -10 0 -4Z"
            className="toile-paper"
          />
          <path
            d="M0 0C8 -4 15 -12 16 -27C11 -16 5 -10 0 -4Z"
            className="toile-paper"
          />
          <path
            d="M0 0C-3 -8 -2 -16 2 -22C1 -14 2 -8 0 0Z"
            className="toile-paper"
          />
        </g>
      </g>
    </g>
  );
}

/** Small five-petal blossom. */
function Blossom({
  x,
  y,
  s = 1,
  d = 0,
}: {
  x: number;
  y: number;
  s?: number;
  d?: number;
}) {
  return (
    <g transform={`translate(${x} ${y}) scale(${s})`}>
      <g className="toile-bloom" style={delay(d)}>
        <g className="toile-twirl" style={gust(d, 4.5 + (d % 1.5))}>
          {Array.from({ length: 5 }, (_, k) => (
            <g key={k} transform={`rotate(${72 * k + 10})`}>
              <path
                d="M0 -2C-7 -5 -8 -13 0 -15C8 -13 7 -5 0 -2Z"
                className="toile-paper"
              />
              <path d="M-2 -5C-3 -8 -3 -10 -2 -12" className="toile-hatch" />
            </g>
          ))}
          <circle r="2.6" className="toile-ink" stroke="none" />
        </g>
      </g>
    </g>
  );
}

/** Serrated rose leaf, hatched on one half, with midrib and veins. */
function leafShape(len: number, w: number, teeth = 6) {
  const steps = teeth * 2;
  const side = (sign: number) =>
    Array.from({ length: steps - 1 }, (_, i) => {
      const t = (i + 1) / steps;
      const hw = w * Math.sin(Math.PI * t ** 0.8) * (i % 2 ? 0.84 : 1);
      return `${n(sign * hw)} ${n(-len * t)}`;
    });
  const outline = `M0 0L${side(1).join("L")}L0 ${-len}L${side(-1).reverse().join("L")}Z`;
  const veins = [0.24, 0.42, 0.6, 0.76].flatMap((t) => {
    const hw = w * Math.sin(Math.PI * (t + 0.1) ** 0.8) * 0.78;
    return [1, -1].map(
      (sg) =>
        `M0 ${n(-len * t)}Q${n(sg * hw * 0.5)} ${n(-len * (t + 0.04))} ${n(sg * hw)} ${n(-len * (t + 0.13))}`,
    );
  });
  const hatch = Array.from({ length: 11 }, (_, i) => {
    const t = 0.1 + i * 0.07;
    const hw = w * Math.sin(Math.PI * (t + 0.05) ** 0.8) * 0.62;
    return `M-1 ${n(-len * t)}L${n(-hw)} ${n(-len * (t + 0.07))}`;
  });
  return {
    outline,
    rib: `M0 0Q${n(w * 0.12)} ${n(-len * 0.5)} 0 ${n(-len * 0.97)}`,
    veins,
    hatch,
  };
}

const BIG_LEAF = leafShape(64, 20);
const SMALL_LEAF = leafShape(40, 13, 5);

function Leaf({
  x,
  y,
  a,
  s = 1,
  d = 0,
  small = false,
}: {
  x: number;
  y: number;
  a: number;
  s?: number;
  d?: number;
  small?: boolean;
}) {
  const leaf = small ? SMALL_LEAF : BIG_LEAF;
  return (
    <g
      transform={`translate(${n(x)} ${n(y)}) rotate(${n(a)}) scale(${s})`}
     
    >
      <g className="toile-grow" style={delay(d)}>
        <g className="toile-flutter" style={gust(d, 5 + (d % 2))}>
          <path d={leaf.outline} className="toile-paper" />
          <path d={leaf.rib} />
          {leaf.veins.map((v) => (
            <path key={v} d={v} className="toile-hatch" />
          ))}
          {leaf.hatch.map((h) => (
            <path key={h} d={h} className="toile-hatch" />
          ))}
        </g>
      </g>
    </g>
  );
}

/** Laurel-like sprig: a thin curved stem bearing pairs of small leaflets. */
function Sprig({
  from,
  c1,
  c2,
  to,
  d = 0,
  pairs = 7,
}: {
  from: P;
  c1: P;
  c2: P;
  to: P;
  d?: number;
  pairs?: number;
}) {
  const pts: P[] = [from, c1, c2, to];
  return (
    <g>
      <path
        d={curve(pts)}
        pathLength={1}
        className="toile-draw"
        style={delay(d, 1.6)}
      />
      {Array.from({ length: pairs }, (_, i) => {
        const t = 0.18 + (i / pairs) * 0.8;
        const { p, angle } = bezier(pts, t);
        const size = 0.95 - (i / pairs) * 0.45;
        return [-38, 38].map((spread) => (
          <g
            key={`${i}${spread}`}
            transform={`translate(${n(p[0])} ${n(p[1])}) rotate(${n(deg(angle) + 90 + spread)}) scale(${n(size)})`}
          >
            <g className="toile-grow" style={delay(d + 0.3 + t * 1.1)}>
              <g className="toile-flutter" style={gust(d + t * 2, 4.5)}>
                <path
                  d="M0 0C-5 -6 -5 -16 0 -22C5 -16 5 -6 0 0Z"
                  className="toile-paper"
                />
                <path d="M0 -2L0 -18" className="toile-hatch" />
              </g>
            </g>
          </g>
        ));
      })}
      <Blossom x={to[0]} y={to[1]} s={0.6} d={d + 1.5} />
    </g>
  );
}

/* ------------------------------------------------------------------ */
/* Composition                                                         */
/* ------------------------------------------------------------------ */

/** Main stem of the left branch: from the bottom center, around the title, up to the top left. */
const STEM: P[] = [
  [776, 748],
  [470, 790],
  [230, 640],
  [318, 236],
];

function Branch() {
  const on = (t: number) => bezier(STEM, t);
  // Leaves along the stem: t, side (1 = outward, -1 = toward the title), size.
  const leaves: [number, number, number][] = [
    [0.08, 1, 0.8],
    [0.2, -1, 0.7],
    [0.47, 1, 1.05],
    [0.58, -1, 0.6],
    [0.68, 1, 0.9],
    [0.8, -1, 0.55],
    [0.88, 1, 0.75],
  ];
  return (
    <g className="toile-sway" style={{ transformOrigin: "776px 748px" }}>
      <path
        d={curve(STEM)}
        pathLength={1}
        className="toile-draw"
        style={delay(0.2, 2.6)}
      />
      {/* secondary stems toward the roses and the outer sprig */}
      <path
        d="M560 768C520 760 498 746 486 728"
        pathLength={1}
        className="toile-draw"
        style={delay(0.9, 0.8)}
      />
      <path
        d="M300 520C280 490 262 470 236 458"
        pathLength={1}
        className="toile-draw"
        style={delay(1.4, 0.8)}
      />

      <Sprig
        from={[296, 470]}
        c1={[214, 430]}
        c2={[176, 360]}
        to={[186, 282]}
        d={1.6}
        pairs={6}
      />
      <Sprig
        from={[672, 768]}
        c1={[640, 812]}
        c2={[580, 836]}
        to={[520, 842]}
        d={1.2}
        pairs={5}
      />

      {leaves.map(([t, side, s]) => {
        const { p, angle } = on(t);
        return (
          <Leaf
            key={t}
            x={p[0]}
            y={p[1]}
            a={deg(angle) + 90 - side * 55}
            s={s}
            small={side < 0}
            d={0.4 + t * 1.8}
          />
        );
      })}
      <Leaf x={520} y={748} a={-150} s={1.1} d={1.1} />
      <Leaf x={448} y={700} a={-80} s={0.9} d={1.3} />
      <Leaf x={300} y={600} a={-110} s={0.85} d={1.5} />

      <Leaf x={420} y={736} a={-120} s={0.75} d={1.2} />
      <Leaf x={540} y={680} a={30} s={0.6} small d={1.4} />
      <Bud x={548} y={700} s={0.7} r={48} d={1.8} />
      <Rose x={482} y={712} s={1.22} r={8} d={1.0} />
      <Rose x={272} y={584} s={0.86} r={-20} d={1.5} />
      <Bud x={236} y={458} s={0.95} r={-62} d={1.9} />
      <Bud x={318} y={236} s={0.85} r={14} d={2.4} />
      <Blossom x={604} y={760} s={0.95} d={1.7} />
      <Blossom x={366} y={660} s={0.8} d={1.9} />
      <Blossom x={330} y={380} s={0.7} d={2.2} />
    </g>
  );
}

/** Light garland above the title: a swag of tiny blossoms and leaflets around a bud. */
function Garland() {
  const swag: P[] = [
    [640, 168],
    [700, 206],
    [900, 206],
    [960, 168],
  ];
  return (
    <g className="toile-swing" style={{ transformOrigin: "800px 150px" }}>
      <path
        d={curve(swag)}
        pathLength={1}
        className="toile-draw"
        style={delay(2, 1.8)}
      />
      {[0.12, 0.3, 0.7, 0.88].map((t, i) => {
        const { p, angle } = bezier(swag, t);
        return (
          <Leaf
            key={t}
            x={p[0]}
            y={p[1]}
            a={deg(angle) + (i < 2 ? 140 : 40)}
            s={0.42}
            small
            d={2.3 + i * 0.15}
          />
        );
      })}
      {[0.2, 0.38, 0.62, 0.8].map((t, i) => {
        const { p } = bezier(swag, t);
        return (
          <Blossom key={t} x={p[0]} y={p[1]} s={0.55} d={2.5 + i * 0.12} />
        );
      })}
      <Bud x={800} y={208} s={0.6} r={180} d={2.8} />
      <Blossom x={800} y={200} s={0.75} d={2.9} />
    </g>
  );
}

/** Bow of tiny blossoms where both branches meet. */
function Knot() {
  return (
    <g>
      <Leaf x={800} y={752} a={-30} s={0.6} small d={1.8} />
      <Leaf x={800} y={752} a={30} s={0.6} small d={1.9} />
      <Bud x={800} y={758} s={0.7} r={0} d={2.1} />
      <Blossom x={778} y={756} s={0.7} d={2.2} />
      <Blossom x={822} y={756} s={0.7} d={2.3} />
    </g>
  );
}

const DRIFT: { x: number; y: number; d: number; dur: number; s: number }[] = [
  { x: 420, y: 120, d: 4, dur: 19, s: 0.9 },
  { x: 1180, y: 90, d: 9, dur: 22, s: 0.7 },
  { x: 980, y: 140, d: 15, dur: 20, s: 0.8 },
  { x: 260, y: 160, d: 21, dur: 24, s: 0.6 },
];

export function ToileFlowers({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 1600 900" className={className} aria-hidden data-toile>
      <g className="toile">
        <Branch />
        <g transform="translate(1600 0) scale(-1 1)">
          <g className="toile-mirror">
            <Branch />
          </g>
        </g>
        <Knot />
        <Garland />
        {DRIFT.map(({ x, y, d, dur, s }) => (
          <g key={x} transform={`translate(${x} ${y}) scale(${s * 0.32})`}>
            <path
              d={PETAL}
              className="toile-drift toile-paper"
              style={delay(d, dur)}
            />
          </g>
        ))}
      </g>
    </svg>
  );
}

/** Where the stem of the spray starts: the end of the subsection trail, in the 1600×900 slide box. */
export const TRAIL_END: P = [958, 790];

/** Single large rose spray (subsection dividers): its stem grows out of the trail and rises to a big rose. */
export function ToileSpray({ className }: { className?: string }) {
  const [x0, y0] = TRAIL_END;
  // Two segments: along the trail line, then up to the rose.
  const foot: P[] = [
    [x0, y0],
    [x0 + 190, y0 + 2],
    [x0 + 330, y0 - 30],
    [x0 + 352, y0 - 190],
  ];
  const rise: P[] = [
    foot[3],
    [x0 + 368, y0 - 320],
    [x0 + 296, y0 - 420],
    [x0 + 340, y0 - 530],
  ];
  const on = (t: number) => bezier(rise, t);
  const leaves: [number, number, number][] = [
    [0.1, -1, 1.5],
    [0.26, 1, 1.4],
    [0.44, -1, 1.6],
    [0.6, 1, 1.25],
    [0.76, -1, 1.1],
  ];
  const [a, b, c, e] = [0.62, 0.42, 0.3, 0.5].map((t) => on(t).p);
  const rose = rise[3];
  return (
    <svg viewBox="0 0 1600 900" className={className} aria-hidden data-toile>
      <g className="toile toile-fine">
        <g
          className="toile-sway"
          style={{ transformOrigin: `${x0}px ${y0}px` }}
        >
          <path
            d={curve(foot)}
            pathLength={1}
            className="toile-draw toile-stem"
            style={delay(0.7, 1.2)}
          />
          <path
            d={curve(rise)}
            pathLength={1}
            className="toile-draw toile-stem"
            style={delay(1.85, 1.3)}
          />
          <path
            d={`M${a}C${a[0] - 50} ${a[1] - 36} ${a[0] - 90} ${a[1] - 90} ${a[0] - 100} ${a[1] - 150}`}
            pathLength={1}
            className="toile-draw"
            style={delay(2.4, 1)}
          />
          <path
            d={`M${b}C${b[0] + 56} ${b[1] - 34} ${b[0] + 108} ${b[1] - 60} ${b[0] + 140} ${b[1] - 116}`}
            pathLength={1}
            className="toile-draw"
            style={delay(2.3, 1)}
          />

          {/* small leaves along the foot of the stem, on the trail */}
          {[0.35, 0.6, 0.82].map((t, i) => {
            const { p, angle } = bezier(foot, t);
            return (
              <Leaf
                key={t}
                x={p[0]}
                y={p[1]}
                a={deg(angle) + 90 + (i % 2 ? 60 : -60)}
                s={0.75 + i * 0.15}
                small
                d={1 + t * 1}
              />
            );
          })}

          <Sprig
            from={c}
            c1={[c[0] + 84, c[1] - 12]}
            c2={[c[0] + 134, c[1] - 56]}
            to={[c[0] + 160, c[1] - 118]}
            d={2.2}
            pairs={7}
          />
          <Sprig
            from={e}
            c1={[e[0] - 72, e[1] + 10]}
            c2={[e[0] - 130, e[1] - 20]}
            to={[e[0] - 156, e[1] - 80]}
            d={2.4}
            pairs={6}
          />

          {leaves.map(([t, side, s]) => {
            const { p, angle } = on(t);
            return (
              <Leaf
                key={t}
                x={p[0]}
                y={p[1]}
                a={deg(angle) + 90 - side * 58}
                s={s}
                d={2 + t * 1.2}
              />
            );
          })}
          <Leaf x={rose[0] - 70} y={rose[1] + 108} a={-78} s={1.4} d={2.9} />
          <Leaf x={rose[0] + 70} y={rose[1] + 104} a={72} s={1.35} d={3} />

          <Bud x={a[0] - 100} y={a[1] - 150} s={1.55} r={-28} d={3} />
          <Bud x={b[0] + 140} y={b[1] - 116} s={1.35} r={34} d={3.1} />
          <Rose x={rose[0]} y={rose[1]} s={2.35} r={-6} d={2.8} />
          <Blossom x={a[0] - 66} y={a[1] - 52} s={1.3} d={3.2} />
          <Blossom x={b[0] + 66} y={b[1] - 14} s={1.15} d={3.3} />
          <Blossom x={rose[0] + 130} y={rose[1] + 24} s={1} d={3.4} />
        </g>
      </g>
    </svg>
  );
}
