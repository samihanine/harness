/**
 * Where a new element goes on a slide, computed by code (the AI cannot see the slide, so it never
 * gives coordinates unless it wants to): the largest free area of the slide or of a frame, the room
 * under or beside a shape — then the element fitted in it, in its proportions, with a margin.
 * Boxes are in EMU (the slide's own unit).
 */
export type Box = { x: number; y: number; w: number; h: number };

export const overlaps = (a: Box, b: Box) => a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
/** `inner` lies in `outer` (a small tolerance: shapes drawn on a frame's edge). */
export const inside = (inner: Box, outer: Box, slack = 0.02) => {
  const dx = outer.w * slack;
  const dy = outer.h * slack;
  return inner.x >= outer.x - dx && inner.y >= outer.y - dy && inner.x + inner.w <= outer.x + outer.w + dx && inner.y + inner.h <= outer.y + outer.h + dy;
};
export const inset = (b: Box, by: number): Box => ({ x: b.x + by, y: b.y + by, w: Math.max(0, b.w - 2 * by), h: Math.max(0, b.h - 2 * by) });

/** A few values spread over a sorted list (keeps the search small on a crowded slide). */
const sample = (values: number[], max: number) => (values.length <= max ? values : Array.from({ length: max }, (_, i) => values[Math.round((i * (values.length - 1)) / (max - 1))]));

/**
 * The largest free rectangle of `bounds` touching none of the obstacles (their edges are the only
 * useful candidates). `aspect` (width / height): the area counted is the one an element of these
 * proportions can use. Undefined when nothing of at least `min` fits.
 */
export function freeRect(bounds: Box, obstacles: Box[], { aspect, minW = bounds.w * 0.12, minH = bounds.h * 0.12 }: { aspect?: number; minW?: number; minH?: number } = {}): Box | undefined {
  const obs = obstacles.filter((o) => overlaps(o, bounds));
  const clamp = (v: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, v));
  const xs = sample([...new Set([bounds.x, bounds.x + bounds.w, ...obs.flatMap((o) => [o.x, o.x + o.w])].map((v) => clamp(v, bounds.x, bounds.x + bounds.w)))].sort((a, b) => a - b), 36);
  const ys = sample([...new Set([bounds.y, bounds.y + bounds.h, ...obs.flatMap((o) => [o.y, o.y + o.h])].map((v) => clamp(v, bounds.y, bounds.y + bounds.h)))].sort((a, b) => a - b), 36);
  let best: Box | undefined;
  let score = 0;
  for (let i = 0; i < xs.length; i++)
    for (let j = i + 1; j < xs.length; j++) {
      const w = xs[j] - xs[i];
      if (w < minW) continue;
      for (let k = 0; k < ys.length; k++)
        for (let l = k + 1; l < ys.length; l++) {
          const h = ys[l] - ys[k];
          if (h < minH) continue;
          const used = aspect ? Math.min(w, h * aspect) * Math.min(h, w / aspect) : w * h;
          if (used <= score) continue;
          const r = { x: xs[i], y: ys[k], w, h };
          if (obs.some((o) => overlaps(o, r))) continue;
          best = r;
          score = used;
        }
    }
  return best;
}

/** An element of given proportions (width / height) fitted in a box, centered. */
export function fit(box: Box, aspect: number): Box {
  const w = Math.min(box.w, box.h * aspect);
  const h = w / aspect;
  return { x: Math.round(box.x + (box.w - w) / 2), y: Math.round(box.y + (box.h - h) / 2), w: Math.round(w), h: Math.round(h) };
}

/** The room under a shape: its width, from its bottom to the next obstacle below (or the bounds' bottom). */
export function below(ref: Box, bounds: Box, obstacles: Box[], gap: number): Box {
  const top = ref.y + ref.h + gap;
  const column = { x: ref.x, y: top, w: ref.w, h: bounds.y + bounds.h - top };
  const next = obstacles.filter((o) => o.y >= top - gap / 2 && overlaps(o, column)).reduce((y, o) => Math.min(y, o.y), bounds.y + bounds.h);
  return { x: ref.x, y: top, w: ref.w, h: Math.max(0, next - gap - top) };
}

/** The room right of a shape: its height, from its right edge to the next obstacle (or the bounds' right edge). */
export function rightOf(ref: Box, bounds: Box, obstacles: Box[], gap: number): Box {
  const left = ref.x + ref.w + gap;
  const row = { x: left, y: ref.y, w: bounds.x + bounds.w - left, h: ref.h };
  const next = obstacles.filter((o) => o.x >= left - gap / 2 && overlaps(o, row)).reduce((x, o) => Math.min(x, o.x), bounds.x + bounds.w);
  return { x: left, y: ref.y, w: Math.max(0, next - gap - left), h: ref.h };
}
