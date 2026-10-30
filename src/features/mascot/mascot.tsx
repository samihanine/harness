/**
 * The mascot: a cartoon hedgehog drawn as in the reference sticker — a big head with a wide white mask,
 * large eyes, a mane of curved quills, thick dark outlines and a white contour around the whole figure
 * (an SVG filter, so every outfit gets it). Parts are free-hand paths in layers: quills, body, legs,
 * arms turning at the shoulder, the head turning at the neck. An outfit fills slots of these layers;
 * a state moves the pivots (motion values looping over keyframes, written as SVG transforms around
 * each part's own pivot, so nothing drifts whatever the size).
 */
import { useEffect, useId, useRef, type ReactNode } from "react";
import { animate, motion, useMotionValue, useMotionValueEvent, useTransform, type MotionValue } from "motion/react";
import { BODY, C, OUTFITS, type Outfit } from "./outfits";

export const STATES = ["idle", "thinking", "talking", "working", "sleeping"] as const;
export type MascotState = (typeof STATES)[number];

/** Keyframes of one part (degrees, px or scale), looped over `duration` seconds. */
type Loop = { values: number[]; duration: number; times?: number[] };
const loop = (values: number[], duration = 3, times?: number[]): Loop => ({ values, duration, times });

/**
 * Each state: how the head, arms, body and eyes move. The right arm's angle: 0 hanging, about -130
 * raised and pointing as in the reference drawing.
 */
const POSES: Record<MascotState, { head: Loop; headY: Loop; armL: Loop; armR: Loop; breath: Loop; blink: Loop }> = {
  idle: { head: loop([0, 3, 0, -3, 0], 6), headY: loop([0]), armL: loop([28, 32, 28], 4), armR: loop([-125, -132, -125], 3), breath: loop([1, 1.02, 1], 3), blink: loop([1, 1, 0.1, 1], 4, [0, 0.9, 0.95, 1]) },
  thinking: { head: loop([-8, -5, -8], 3), headY: loop([0]), armL: loop([28]), armR: loop([-150, -145, -150], 3), breath: loop([1, 1.015, 1], 3), blink: loop([1, 1, 0.1, 1], 5, [0, 0.92, 0.96, 1]) },
  talking: { head: loop([0, -3, 2, -2, 0], 1.6), headY: loop([0, -2, 0, -1.5, 0], 1.6), armL: loop([28, 34, 28], 1.6), armR: loop([-70, -130, -95, -135, -70], 2.4), breath: loop([1, 1.02, 1], 1.6), blink: loop([1, 1, 0.1, 1], 3.5, [0, 0.9, 0.95, 1]) },
  working: { head: loop([4, 7, 4], 0.7), headY: loop([0, 1.5, 0], 0.7), armL: loop([-20, -40, -20], 0.7), armR: loop([20, 45, 20], 0.7), breath: loop([1, 1.02, 1], 0.7), blink: loop([0.75]) },
  sleeping: { head: loop([10, 13, 10], 4), headY: loop([4, 6, 4], 4), armL: loop([8]), armR: loop([-8]), breath: loop([1, 1.045, 1], 4), blink: loop([0]) },
};

/** A motion value playing a loop (or moving to a single value), started again when the loop changes. */
function useLoop({ values, duration, times }: Loop) {
  const v = useMotionValue(values[0]);
  const key = `${values.join()}|${duration}|${times?.join()}`;
  useEffect(() => {
    const c = values.length > 1 ? animate(v, values, { duration, times, repeat: Infinity, ease: "easeInOut" }) : animate(v, values[0], { duration: 0.5 });
    return () => c.stop();
  }, [key]); // eslint-disable-line react-hooks/exhaustive-deps
  return v;
}

/** `rotate(a cx cy)` of a part, from a motion value of its angle. */
const useTurn = (angle: MotionValue<number>, cx: number, cy: number) => useTransform(angle, (a) => `rotate(${a} ${cx} ${cy})`);
/** A vertical squash around a point (breathing, blinking). */
const useSquash = (s: MotionValue<number>, cx: number, cy: number) => useTransform(s, (k) => `translate(${cx} ${cy}) scale(1 ${k}) translate(${-cx} ${-cy})`);

/** The drawing's outline: thick and round, as in the sticker. */
export const S = { stroke: C.line, strokeWidth: 3.6, strokeLinejoin: "round" as const, strokeLinecap: "round" as const };

/** A group whose SVG `transform` attribute follows a motion value (written on the element at each frame). */
function Moving({ transform, children }: { transform: MotionValue<string>; children: ReactNode }) {
  const ref = useRef<SVGGElement>(null);
  useMotionValueEvent(transform, "change", (v) => ref.current?.setAttribute("transform", v));
  return (
    <g ref={ref} transform={transform.get()}>
      {children}
    </g>
  );
}

/* --------------------------------- Drawing --------------------------------- */

/** The head's center: the mane is drawn around it. */
const HEAD = { x: 100, y: 84 };

/**
 * The mane: curved quills around the head and down the back, each a dark blade bent the same way with
 * its lit face — longer at the top and on the back (the left of the drawing), shorter low on the right.
 */
const QUILLS = (() => {
  const pt = (deg: number, r: number) => {
    const a = (deg * Math.PI) / 180;
    return [HEAD.x + r * Math.cos(a), HEAD.y + 4 + r * Math.sin(a)];
  };
  const f = (p: number[]) => `${p[0].toFixed(1)} ${p[1].toFixed(1)}`;
  const n = 16;
  const [from, to] = [118, 412];
  return Array.from({ length: n }, (_, i) => {
    const mid = from + ((to - from) * (i + 0.5)) / n;
    const half = ((to - from) / n) * 0.95;
    // Longest on the back (around 180°–250°), shortest low on the right.
    const back = Math.max(0, Math.cos(((mid - 205) * Math.PI) / 180));
    const length = 72 + 18 * back + (i % 2 ? 0 : 5) - (mid > 380 ? 10 : 0);
    const bend = 13;
    const tip = pt(mid + bend, length);
    const b0 = pt(mid - half, 44);
    const b1 = pt(mid + half, 44);
    // The blade curves: its leading edge bulges outwards, its trailing edge comes back in.
    const c0 = pt(mid - half + bend * 0.2, length * 0.78);
    const c1 = pt(mid + half + bend * 0.9, length * 0.62);
    const lit0 = pt(mid - half * 0.35, 50);
    const litC = pt(mid + bend * 0.35, length * 0.74);
    const litTip = pt(mid + bend * 0.85, length - 9);
    return {
      blade: `M${f(b0)} Q${f(c0)} ${f(tip)} Q${f(c1)} ${f(b1)} Z`,
      light: `M${f(lit0)} Q${f(litC)} ${f(litTip)} Q${f(pt(mid + half * 0.6, length * 0.6))} ${f(pt(mid + half * 0.55, 50))} Z`,
    };
  });
})();

function Eye({ cx, blink, state }: { cx: number; blink: MotionValue<number>; state: MascotState }) {
  const transform = useSquash(blink, cx, 84);
  if (state === "sleeping") return <path d={`M${cx - 8} 84 Q${cx} 91 ${cx + 8} 84`} fill="none" {...S} strokeWidth={3.2} />;
  return (
    <Moving transform={transform}>
      <ellipse cx={cx} cy={85} rx={9.6} ry={12} fill={C.line} />
      {/* Looking up while thinking. */}
      <motion.g animate={{ y: state === "thinking" ? -2.5 : 0, x: state === "thinking" ? 1 : 0 }}>
        <ellipse cx={cx + 3} cy={79.5} rx={3.6} ry={4} fill={C.white} />
        <circle cx={cx - 3.2} cy={91} r={1.7} fill={C.white} />
      </motion.g>
    </Moving>
  );
}

/** The open smile of the drawing (a dark D with its tongue), or the state's mouth. */
function Mouth({ state }: { state: MascotState }) {
  if (state === "talking")
    return <motion.ellipse cx={101} cy={105} rx={8} fill={C.mouth} stroke={C.line} strokeWidth={2.6} animate={{ ry: [2.5, 7, 3.5, 8, 3, 6, 2.5] }} transition={{ duration: 1.2, repeat: Infinity }} />;
  if (state === "thinking") return <ellipse cx={106} cy={105} rx={3.6} ry={3.1} fill={C.mouth} stroke={C.line} strokeWidth={2.6} />;
  if (state === "sleeping") return <motion.ellipse cx={101} cy={105} rx={3.6} fill={C.mouth} stroke={C.line} strokeWidth={2.2} animate={{ ry: [1.8, 3.2, 1.8] }} transition={{ duration: 4, repeat: Infinity }} />;
  if (state === "working") return <path d="M92 103 Q101 108 110 102" fill="none" {...S} strokeWidth={3} />;
  return (
    <>
      <path d="M88.5 99 Q101 102 113.5 98.5 Q113 113 101 114 Q89.5 113.5 88.5 99 Z" fill={C.mouth} {...S} strokeWidth={2.8} />
      <path d="M93.5 109.5 Q101 104.5 109 109 Q106 113.2 101 113.4 Q96 113.2 93.5 109.5 Z" fill={C.tongue} />
    </>
  );
}

/** What floats around the mascot: a thought bubble, Zs while sleeping. */
function Extras({ state }: { state: MascotState }) {
  if (state === "thinking")
    return (
      <>
        {[
          [156, 40, 4],
          [166, 26, 6],
          [181, 10, 9],
        ].map(([x, y, r], i) => (
          <motion.circle key={i} cx={x} cy={y} r={r} fill={C.white} {...S} strokeWidth={2.5} animate={{ opacity: [0.2, 1, 0.2] }} transition={{ duration: 1.5, repeat: Infinity, delay: i * 0.3 }} />
        ))}
      </>
    );
  if (state === "sleeping")
    return (
      <>
        {[0, 1, 2].map((i) => (
          <motion.text key={i} x={146 + i * 9} y={40} fontSize={13 + i * 4} fontWeight={800} fill={C.accent} stroke={C.line} strokeWidth={1} animate={{ y: [0, -24], opacity: [0, 1, 0] }} transition={{ duration: 3, repeat: Infinity, delay: i }}>
            Z
          </motion.text>
        ))}
      </>
    );
  return null;
}

/** An arm hanging from its shoulder (x, 126), the mitten at its end; `sleeve`: the outfit's color. */
function Arm({ x, side, transform, sleeve, prop }: { x: number; side: -1 | 1; transform: MotionValue<string>; sleeve?: string; prop?: ReactNode }) {
  return (
    <Moving transform={transform}>
      <path d={`M${x - 10} 124 Q${x - 12} 142 ${x - 9} 157 L${x + 9} 157 Q${x + 12} 140 ${x + 10} 124 Z`} fill={sleeve ?? C.suit} {...S} />
      {/* Props are drawn for a hand at x 126: moved to this one. */}
      <g transform={`translate(${x - 126} 0)`}>{prop}</g>
      {/* The mitten: a round hand, its thumb and two finger lines. */}
      <path d={`M${x - 11} 157 Q${x - 12} 171 ${x} 172 Q${x + 12} 171 ${x + 11} 157 Q${x} 150 ${x - 11} 157 Z`} fill={C.suit} {...S} />
      <path d={`M${x + side * 9} 158 Q${x + side * 15} 156 ${x + side * 13} 163`} fill={C.suit} {...S} strokeWidth={3} />
      <path d={`M${x - 4} 164 L${x - 4} 169 M${x + 2} 164 L${x + 2} 169`} stroke={C.line} strokeWidth={1.8} strokeLinecap="round" />
    </Moving>
  );
}

/**
 * The mascot dressed in `outfit`, playing `state` (idle, thinking, talking, working, sleeping).
 * `size`: its width in px (the drawing keeps its proportions); `still`: no movement (small icons).
 */
export function Mascot({ outfit = "diver", state = "idle", size = 96, still, className, title }: { outfit?: Outfit; state?: MascotState; size?: number; still?: boolean; className?: string; title?: string }) {
  const pose = POSES[state];
  const keep = (l: Loop) => (still ? loop([l.values[0]]) : l);
  const slots = OUTFITS[outfit] ?? OUTFITS.diver;
  const head = useTransform([useLoop(keep(pose.head)), useLoop(keep(pose.headY))], ([a, y]) => `translate(0 ${y}) rotate(${a} 100 122)`);
  const armL = useTurn(useLoop(keep(pose.armL)), 69, 128);
  const armR = useTurn(useLoop(keep(pose.armR)), 131, 128);
  const breath = useSquash(useLoop(keep(pose.breath)), 100, 186);
  const blink = useLoop(keep(pose.blink));
  const sticker = `sticker-${useId().replace(/:/g, "")}`;
  return (
    <svg viewBox="4 -6 192 224" width={size} height={(size * 224) / 192} className={className} role="img" aria-label={title ?? `Mascot (${outfit}, ${state})`} overflow="visible">
      {title && <title>{title}</title>}
      <defs>
        {/* The sticker's white contour around the whole figure. */}
        <filter id={sticker} x="-15%" y="-15%" width="130%" height="130%">
          <feMorphology in="SourceAlpha" operator="dilate" radius={4} result="grown" />
          <feFlood floodColor={C.white} />
          <feComposite in2="grown" operator="in" result="contour" />
          <feMerge>
            <feMergeNode in="contour" />
            <feMergeNode in="SourceGraphic" />
          </feMerge>
        </filter>
      </defs>
      <g filter={`url(#${sticker})`}>
        {/* The mane moves with the head (it grows around it: no gap when the head turns or nods). */}
        <Moving transform={head}>
          {QUILLS.map((q, i) => (
            <g key={i}>
              <path d={q.blade} fill={C.spikes} {...S} strokeWidth={3.2} />
              <path d={q.light} fill={C.spikeLight} />
            </g>
          ))}
          {"hood" in slots && slots.hood}
        </Moving>
        <Moving transform={breath}>{"back" in slots && slots.back}</Moving>
        {/* Legs, then the feet (or the outfit's). */}
        <path d="M80 172 L79 192 Q88 197 97 192 L97 174 Z" fill={C.suit} {...S} />
        <path d="M103 174 L103 192 Q112 197 121 192 L120 172 Z" fill={C.suit} {...S} />
        {"feet" in slots ? slots.feet : null}
        <Moving transform={breath}>
          {/* The body: a broad pear with the white belly, its right side in shade. */}
          <path d={BODY} fill={C.suit} {...S} />
          <path d="M125 130 Q138 154 130 175 Q137 154 125 130 Z" fill={C.suitShade} />
          <path d="M83 142 Q81 126 100 126 Q119 126 117 142 Q119 171 100 173 Q81 171 83 142 Z" fill={C.white} stroke={C.line} strokeWidth={2.4} />
          {"torso" in slots && slots.torso}
        </Moving>
        <Arm x={69} side={-1} transform={armL} sleeve={"sleeve" in slots ? (slots.sleeve as string) : undefined} />
        <Moving transform={head}>
          {/* Ears on the sides, pink inside. */}
          <circle cx={50} cy={68} r={11} fill={C.fur} {...S} />
          <circle cx={150} cy={68} r={11} fill={C.fur} {...S} />
          <circle cx={49} cy={68} r={6} fill={C.pink} />
          <circle cx={151} cy={68} r={6} fill={C.pink} />
          {/* The head: wide, a little squared at the cheeks. */}
          <path d="M100 42 C130 42 150 56 151 82 C152 108 132 124 100 124 C68 124 48 108 49 82 C50 56 70 42 100 42 Z" fill={C.fur} {...S} />
          <path d="M136 58 Q152 84 138 112 Q147 84 136 58 Z" fill={C.furShade} />
          {/* The white mask: around the eyes and down to the chin, the fur's point between the eyes. */}
          <path d="M58 92 C55 72 68 64 82 66 C91 67 96 71 100 77 C104 71 109 67 118 66 C132 64 145 72 142 92 C140 112 124 121 100 121 C76 121 60 112 58 92 Z" fill={C.white} stroke={C.line} strokeWidth={2.4} />
          <ellipse cx={68} cy={100} rx={6.5} ry={4.5} fill={C.pink} opacity={0.85} />
          <ellipse cx={132} cy={100} rx={6.5} ry={4.5} fill={C.pink} opacity={0.85} />
          <Eye cx={82} blink={blink} state={state} />
          <Eye cx={119} blink={blink} state={state} />
          <ellipse cx={101} cy={96} rx={7} ry={4.8} fill={C.line} />
          <ellipse cx={98.5} cy={94.4} rx={2.4} ry={1.3} fill={C.white} opacity={0.85} />
          <Mouth state={state} />
          {"eyes" in slots && slots.eyes}
          {"face" in slots && slots.face}
          {"headTop" in slots && slots.headTop}
        </Moving>
        <Arm x={131} side={1} transform={armR} sleeve={"sleeve" in slots ? (slots.sleeve as string) : undefined} prop={"prop" in slots ? slots.prop : undefined} />
      </g>
      {!still && <Extras state={state} />}
    </svg>
  );
}
