/**
 * Facts computed from a visual's exported data, given to the AI with "Explain" so it narrates figures
 * instead of computing them: per value column the total, leader and share, the last point vs the
 * average, the biggest jump between two points, the trend, and outliers (far from the others).
 */

/** CSV as Power BI exports it (quoted fields, "" inside quotes). */
export function parseCsv(csv: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  const text = csv.replace(/\r/g, "").replace(/^﻿/, "");
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') (cell += '"'), i++;
      else if (ch === '"') quoted = false;
      else cell += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ",") row.push(cell), (cell = "");
    else if (ch === "\n") row.push(cell), rows.push(row), (row = []), (cell = "");
    else cell += ch;
  }
  if (cell || row.length) row.push(cell), rows.push(row);
  return rows.filter((r) => r.some((c) => c.trim()));
}

const number = (s: string) => {
  const t = s.trim().replace(/[\s ]/g, "");
  return /^-?\d+(\.\d+)?(e-?\d+)?%?$/i.test(t) ? Number(t.replace("%", "")) : undefined;
};
export const round = (n: number) => {
  const a = Math.abs(n);
  const r = (x: number, s: string) => `${+(n / x).toFixed(Math.abs(n / x) >= 100 ? 0 : 1)}${s}`;
  return a >= 1e9 ? r(1e9, "B") : a >= 1e6 ? r(1e6, "M") : a >= 1e4 ? r(1e3, "K") : `${+n.toFixed(a >= 100 ? 0 : 2)}`;
};
const pct = (x: number) => `${x >= 0 ? "+" : ""}${Math.round(x * 100)} %`;

/** The facts of a CSV export (first column(s) = labels, numeric columns = values); empty when there is nothing to say. */
export function dataFacts(csv: string, max = 12) {
  const [head, ...rows] = parseCsv(csv);
  if (!head || rows.length < 2) return "";
  const numeric = head.map((_, i) => rows.filter((r) => r[i]?.trim()).every((r) => number(r[i]) !== undefined) && rows.some((r) => r[i]?.trim()));
  const labels = head.map((_, i) => i).filter((i) => !numeric[i]);
  const label = (r: string[]) => labels.map((i) => r[i]).join(" / ") || "?";
  const facts: string[] = [];
  for (const [i, name] of head.entries()) {
    if (!numeric[i]) continue;
    const points = rows.map((r) => ({ label: label(r), v: number(r[i] ?? "") })).filter((p): p is { label: string; v: number } => p.v !== undefined);
    if (points.length < 2) continue;
    const values = points.map((p) => p.v);
    const total = values.reduce((a, b) => a + b, 0);
    const mean = total / values.length;
    const sd = Math.sqrt(values.reduce((a, v) => a + (v - mean) ** 2, 0) / values.length);
    const sorted = [...points].sort((a, b) => b.v - a.v);
    const [top, bottom] = [sorted[0], sorted.at(-1)!];
    const line = [`[${name}] over ${points.length} rows: total ${round(total)}, average ${round(mean)}`];
    if (total > 0 && values.every((v) => v >= 0)) {
      line.push(`leader ${top.label} ${round(top.v)} (${Math.round((top.v / total) * 100)} % of the total)`);
      const top3 = sorted.slice(0, 3).reduce((a, p) => a + p.v, 0);
      if (points.length > 4) line.push(`top 3 = ${Math.round((top3 / total) * 100)} %`);
    } else line.push(`highest ${top.label} ${round(top.v)}`);
    line.push(`lowest ${bottom.label} ${round(bottom.v)}`);
    // In the export's order (a time axis is exported in order): last point, jumps, trend.
    const last = points.at(-1)!;
    if (mean) line.push(`last (${last.label}) ${round(last.v)} = ${pct(last.v / mean - 1)} vs the average`);
    const jumps = points.slice(1).map((p, k) => ({ from: points[k], to: p, d: p.v - points[k].v }));
    const jump = jumps.sort((a, b) => Math.abs(b.d) - Math.abs(a.d))[0];
    if (jump && sd && Math.abs(jump.d) > sd) line.push(`biggest jump ${jump.from.label} → ${jump.to.label}: ${jump.d > 0 ? "+" : ""}${round(jump.d)}${jump.from.v ? ` (${pct(jump.d / jump.from.v)})` : ""}`);
    if (points.length >= 4) {
      const n = points.length;
      const xm = (n - 1) / 2;
      const slope = values.reduce((a, v, k) => a + (k - xm) * (v - mean), 0) / values.reduce((a, _, k) => a + (k - xm) ** 2, 0);
      const change = (slope * (n - 1)) / (Math.abs(mean) || 1);
      if (Math.abs(change) > 0.1) line.push(`in row order the values go ${change > 0 ? "up" : "down"} (${pct(change)} first → last on the trend line)`);
    }
    const outliers = points.length >= 6 && sd ? points.filter((p) => Math.abs(p.v - mean) > 2 * sd) : [];
    if (outliers.length) line.push(`outliers (beyond 2 standard deviations): ${outliers.slice(0, 3).map((p) => `${p.label} ${round(p.v)}`).join(", ")}`);
    facts.push(`- ${line.join("; ")}`);
    if (facts.length >= max) break;
  }
  return facts.join("\n");
}
