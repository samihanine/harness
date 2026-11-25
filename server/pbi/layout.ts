/**
 * The layout engine of the Builder: the AI gives rows of visuals (heights in px or shared, widths as
 * relative weights) and this code computes every position on the page — margins, even gaps, inside
 * the page, no overlap. Free positions (place_visuals) for what rows cannot say: layers (cards over a frame
 * shape), headers, exact spots — snapped to a grid and checked the same way.
 */

export type LayoutRow = { height?: number; items: { visual: string; width?: number }[] };
export type Box = { visual: string; x: number; y: number; width: number; height: number };



export function layoutRows(rows: LayoutRow[], page: { width: number; height: number }, { margin = 20, gap = 16, top = 0 } = {}): Box[] {
  const used = rows.filter((r) => r.items?.length);
  if (!used.length) throw new Error('Give "rows": each with "items" (the visuals of the row, left to right).');
  const seen = new Set<string>();
  for (const v of used.flatMap((r) => r.items.map((i) => i.visual))) {
    if (seen.has(v)) throw new Error(`Visual ${v} is in two places: give each visual once.`);
    seen.add(v);
  }
  const innerH = page.height - 2 * margin - top - gap * (used.length - 1);
  const fixed = used.reduce((n, r) => n + (r.height ?? 0), 0);
  const free = used.filter((r) => !r.height).length;
  const share = free ? Math.floor((innerH - fixed) / free) : 0;
  if (free && share < 60) throw new Error(`The rows with a height take ${fixed} px of ${innerH} px: the others would be ${share} px high. Lower some heights or make the page higher (resize_page).`);
  if (!free && fixed > innerH) throw new Error(`The rows take ${fixed} px but only ${innerH} px fit (page ${page.height} px, margins ${margin}, gaps ${gap}). Lower some heights, or make the page higher (resize_page).`);
  const boxes: Box[] = [];
  let y = margin + top;
  for (const row of used) {
    const height = row.height ?? share;
    const weights = row.items.map((i) => (i.width && i.width > 0 ? i.width : 1));
    const total = weights.reduce((a, b) => a + b, 0);
    const innerW = page.width - 2 * margin - gap * (row.items.length - 1);
    let x = margin;
    row.items.forEach((item, k) => {
      // The last one takes what rounding left: the row ends exactly at the margin.
      const width = k === row.items.length - 1 ? page.width - margin - x : Math.floor((innerW * weights[k]) / total);
      boxes.push({ visual: item.visual, x, y, width, height });
      x += width + gap;
    });
    y += height + gap;
  }
  return boxes;
}

/** Rows laid out inside an area of the page (e.g. cards inside a frame): positions offset to it. */
export function layoutRowsIn(rows: LayoutRow[], area: { x: number; y: number; width: number; height: number }, o: { margin?: number; gap?: number }) {
  return layoutRows(rows, area, o).map((b) => ({ ...b, x: b.x + area.x, y: b.y + area.y }));
}

/** Decorations: what may lie under other visuals (a frame, a band, a logo, a title zone). */
const LAYER = ["shape", "image", "basicShape", "textbox"];

/** A free position snapped to the grid and kept inside the page. */
export function snapBox<T extends { x: number; y: number; width: number; height: number }>(b: T, page: { width: number; height: number }, grid = 4): T {
  const r = (n: number) => Math.round(n / grid) * grid;
  const width = Math.min(page.width, Math.max(grid, r(b.width)));
  const height = Math.min(page.height, Math.max(grid, r(b.height)));
  return { ...b, width, height, x: Math.min(page.width - width, Math.max(0, r(b.x))), y: Math.min(page.height - height, Math.max(0, r(b.y))) };
}

/**
 * Visuals overlapping or out of the page (positions as the definition holds them). Layers are allowed: a
 * visual over a decoration (shape, image, text box) drawn under it (lower z).
 */
export function layoutProblems(visuals: { name: string; x: number; y: number; width: number; height: number; hidden?: boolean; type?: string; z?: number }[], page: { width: number; height: number }) {
  const shown = visuals.filter((v) => !v.hidden);
  const problems: string[] = [];
  for (const v of shown) if (v.x < 0 || v.y < 0 || v.x + v.width > page.width + 1 || v.y + v.height > page.height + 1) problems.push(`${v.name} goes out of the page (${page.width}×${page.height})`);
  for (let i = 0; i < shown.length; i++)
    for (let j = i + 1; j < shown.length; j++) {
      const [a, b] = [shown[i], shown[j]];
      const w = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x);
      const h = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y);
      const [under] = [a, b].sort((x, y) => (x.z ?? 0) - (y.z ?? 0));
      if (w > 4 && h > 4 && !LAYER.includes(under.type ?? "")) problems.push(`${a.name} and ${b.name} overlap`);
    }
  return problems;
}
