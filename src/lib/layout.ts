/**
 * The layout engine of the Builder: the AI gives rows of visuals (heights in px or shared, widths as
 * relative weights) and this code computes every position on the page — margins, even gaps, inside
 * the page, no overlap. Free positions stay possible (create_visual / move_visual) for the rare
 * layouts rows cannot say.
 */

export type LayoutRow = { height?: number; items: { visual: string; width?: number }[] };
export type Box = { visual: string; x: number; y: number; width: number; height: number };

export const LAYOUT_HELP = `layout_page places the visuals of a page in rows (the app computes x, y, width, height: margins, even gaps, inside the page). Each row: "items" (visuals, left to right; "width" = relative weight, default 1, e.g. 2 and 1 = two thirds / one third) and "height" in px (rows without height share the rest of the page). Typical: a KPI row (height 110), then charts. A visual left out keeps its place.`;

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
  if (free && share < 60) throw new Error(`The rows with a height take ${fixed} px of ${innerH} px: the others would be ${share} px high. Lower some heights or make the page higher.`);
  if (!free && fixed > innerH) throw new Error(`The rows take ${fixed} px but only ${innerH} px fit (page ${page.height} px, margins ${margin}, gaps ${gap}). Lower some heights, or make the page higher (update_page).`);
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

/** Visuals overlapping or out of the page (positions as the definition holds them). */
export function layoutProblems(visuals: { name: string; x: number; y: number; width: number; height: number; hidden?: boolean }[], page: { width: number; height: number }) {
  const shown = visuals.filter((v) => !v.hidden);
  const problems: string[] = [];
  for (const v of shown) if (v.x < 0 || v.y < 0 || v.x + v.width > page.width + 1 || v.y + v.height > page.height + 1) problems.push(`${v.name} goes out of the page (${page.width}×${page.height})`);
  for (let i = 0; i < shown.length; i++)
    for (let j = i + 1; j < shown.length; j++) {
      const [a, b] = [shown[i], shown[j]];
      const w = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x);
      const h = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y);
      if (w > 4 && h > 4) problems.push(`${a.name} and ${b.name} overlap`);
    }
  return problems;
}
