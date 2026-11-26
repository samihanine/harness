/**
 * What a sheet of the workbook shows, for the screen and for the AI's eyes: the texts Excel displays (computed
 * by Excel, CUBE formulas included), the formats of the file (fonts, fills, borders, widths, heights, merged
 * cells) and its conditional formats worked out on the values — the first rows and columns of the sheet. The
 * same HTML is drawn in the app and pictured for the AI.
 */
import ExcelJS from "exceljs";
import { closeSession, excel, sheetPath, workbookBytes } from "./excel";

export type CellStyle = { b?: 1; i?: 1; u?: 1; sz?: number; color?: string; fill?: string; h?: string; v?: string; wrap?: 1; bt?: string; br?: string; bb?: string; bl?: string; indent?: number };
export type SheetPreview = { sheets: string[]; sheet: string; widths: number[]; heights: (number | undefined)[]; cells: (string | null)[][]; styles: (CellStyle | undefined)[][]; merges: string[]; total: { rows: number; cols: number }; address?: string };

const MAX_ROWS = 150;
const MAX_COLS = 30;

const argb = (c?: { argb?: string; theme?: number }) => (c?.argb && c.argb.length >= 6 ? `#${c.argb.slice(-6)}` : undefined);
const border = (b?: { style?: string; color?: { argb?: string } }) => (b?.style ? `${b.style === "thick" ? 2 : b.style === "medium" ? 1.5 : 1}px ${b.style === "dashed" ? "dashed" : b.style === "dotted" ? "dotted" : b.style === "double" ? "double" : "solid"} ${argb(b.color) ?? "#000"}` : undefined);

function styleOf(cell: ExcelJS.Cell): CellStyle | undefined {
  const f = cell.font ?? {};
  const fill = cell.fill?.type === "pattern" && cell.fill.pattern === "solid" ? argb(cell.fill.fgColor) : undefined;
  const a = cell.alignment ?? {};
  const bd = cell.border ?? {};
  const s: CellStyle = {
    ...(f.bold && { b: 1 as const }),
    ...(f.italic && { i: 1 as const }),
    ...(f.underline && { u: 1 as const }),
    ...(f.size && f.size !== 11 && { sz: f.size }),
    ...(argb(f.color) && argb(f.color) !== "#000000" && { color: argb(f.color) }),
    ...(fill && { fill }),
    ...(a.horizontal && { h: a.horizontal }),
    ...(a.vertical && { v: a.vertical }),
    ...(a.wrapText && { wrap: 1 as const }),
    ...(a.indent && { indent: a.indent }),
    ...(border(bd.top) && { bt: border(bd.top) }),
    ...(border(bd.right) && { br: border(bd.right) }),
    ...(border(bd.bottom) && { bb: border(bd.bottom) }),
    ...(border(bd.left) && { bl: border(bd.left) }),
  };
  return Object.keys(s).length ? s : undefined;
}

const colIndex = (letters: string) => letters.split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const parseRef = (ref: string) => {
  const m = ref.match(/^\$?([A-Z]+)\$?(\d+)$/i)!;
  return { c: colIndex(m[1].toUpperCase()), r: Number(m[2]) };
};
const inRange = (ref: string, r: number, c: number) => ref.split(/\s+/).some((part) => {
  const [a, b] = part.split(":");
  const s = parseRef(a);
  const e = b ? parseRef(b) : s;
  return r >= s.r && r <= e.r && c >= s.c && c <= e.c;
});

/** Conditional formats of the file applied to the shown values (the common rules: comparisons, text, color scales). */
function applyConditional(ws: ExcelJS.Worksheet, cells: (string | null)[][], values: unknown[][], styles: (CellStyle | undefined)[][]) {
  const list = ((ws as unknown as { conditionalFormattings?: { ref: string; rules: any[] }[] }).conditionalFormattings ?? []);
  const num = (v: unknown) => (typeof v === "number" ? v : typeof v === "string" && v.trim() !== "" && !Number.isNaN(Number(v)) ? Number(v) : undefined);
  for (const cf of list) {
    for (const rule of cf.rules ?? []) {
      const fillOf = (st: any) => argb(st?.fill?.bgColor) ?? argb(st?.fill?.fgColor);
      const fontOf = (st: any) => argb(st?.font?.color);
      if (rule.type === "colorScale") {
        const colors: string[] = (rule.color ?? []).map((c: any) => argb(c) ?? "#FFFFFF");
        const pts: number[] = [];
        values.forEach((row, r) => row.forEach((v, c) => inRange(cf.ref, r + 1, c + 1) && num(v) !== undefined && pts.push(num(v)!)));
        if (!pts.length || colors.length < 2) continue;
        const min = Math.min(...pts), max = Math.max(...pts);
        const mix = (a: string, b: string, t: number) => `#${[1, 3, 5].map((i) => Math.round(parseInt(a.slice(i, i + 2), 16) * (1 - t) + parseInt(b.slice(i, i + 2), 16) * t).toString(16).padStart(2, "0")).join("")}`;
        values.forEach((row, r) =>
          row.forEach((v, c) => {
            const n = num(v);
            if (!inRange(cf.ref, r + 1, c + 1) || n === undefined) return;
            const t = max === min ? 0.5 : (n - min) / (max - min);
            const fill = colors.length === 3 ? (t < 0.5 ? mix(colors[0], colors[1], t * 2) : mix(colors[1], colors[2], (t - 0.5) * 2)) : mix(colors[0], colors[1], t);
            styles[r][c] = { ...(styles[r][c] ?? {}), fill };
          }),
        );
        continue;
      }
      const test = (v: unknown, text: string | null): boolean => {
        const n = num(v);
        const f = (rule.formulae ?? []).map((x: string) => num(String(x).replace(/^"|"$/g, "")) ?? String(x).replace(/^"|"$/g, ""));
        if (rule.type === "cellIs") {
          if (n === undefined && typeof f[0] === "number") return false;
          const x: any = n ?? text;
          return { greaterThan: x > f[0], lessThan: x < f[0], greaterThanOrEqual: x >= f[0], lessThanOrEqual: x <= f[0], equal: x == f[0], notEqual: x != f[0], between: x >= f[0] && x <= f[1], notBetween: x < f[0] || x > f[1] }[rule.operator as string] ?? false;
        }
        if (rule.type === "containsText") return !!text && text.toLowerCase().includes(String(rule.text ?? "").toLowerCase());
        return false;
      };
      values.forEach((row, r) =>
        row.forEach((v, c) => {
          if (!inRange(cf.ref, r + 1, c + 1) || !test(v, cells[r][c])) return;
          styles[r][c] = { ...(styles[r][c] ?? {}), ...(fillOf(rule.style) && { fill: fillOf(rule.style) }), ...(fontOf(rule.style) && { color: fontOf(rule.style) }), ...(rule.style?.font?.bold && { b: 1 as const }) };
        }),
      );
    }
  }
}

/** A sheet as shown: texts from Excel (its session: what the AI just changed), formats from the file. */
export async function sheetPreview(itemId: string, sheet?: string): Promise<SheetPreview> {
  const sheets = ((await excel(itemId, "/worksheets?$select=name,position")).value as { name: string; position: number }[]).sort((a, b) => a.position - b.position).map((s) => s.name);
  const name = sheet && sheets.includes(sheet) ? sheet : sheets[0];
  const used = await excel(itemId, `${sheetPath(name)}/usedRange(valuesOnly=false)?$select=address,rowCount,columnCount`);
  const rows = Math.min(MAX_ROWS, used.rowCount ?? 0);
  const cols = Math.min(MAX_COLS, used.columnCount ?? 0);
  // From A1 (the sheet as it looks), to the used range's end within the limits.
  const end = used.address?.split("!")[1]?.split(":")[1];
  const last = end ? parseRef(end.replace(/\$/g, "")) : { r: 1, c: 1 };
  const nr = Math.min(MAX_ROWS, last.r);
  const nc = Math.min(MAX_COLS, last.c);
  const letters = (n: number): string => (n > 26 ? letters(Math.floor((n - 1) / 26)) : "") + String.fromCharCode(65 + ((n - 1) % 26));
  const range = await excel(itemId, `${sheetPath(name)}/range(address='A1:${letters(Math.max(1, nc))}${Math.max(1, nr)}')?$select=text,values`);
  const cells: (string | null)[][] = (range.text as string[][]).map((r) => r.map((t) => (t === "" ? null : t)));
  const values: unknown[][] = range.values;
  // The file's formats (as Excel saved it: the session's last changes flushed first).
  await closeSession(itemId);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load((await workbookBytes(itemId)) as never);
  const ws = wb.getWorksheet(name);
  const styles: (CellStyle | undefined)[][] = cells.map((row, r) => row.map((_, c) => (ws ? styleOf(ws.getCell(r + 1, c + 1)) : undefined)));
  if (ws) applyConditional(ws, cells, values, styles);
  const widths = Array.from({ length: Math.max(1, nc) }, (_, c) => Math.round(((ws?.getColumn(c + 1).width ?? 8.43) * 7 + 5)));
  const heights = Array.from({ length: Math.max(1, nr) }, (_, r) => (ws?.getRow(r + 1).height ? Math.round(ws.getRow(r + 1).height * 1.33) : undefined));
  const merges = ((ws as unknown as { model?: { merges?: string[] } })?.model?.merges ?? []).filter((m) => {
    const [a] = m.split(":");
    const p = parseRef(a);
    return p.r <= nr && p.c <= nc;
  });
  void rows;
  void cols;
  return { sheets, sheet: name, widths, heights, cells, styles, merges, total: { rows: last.r, cols: last.c }, address: used.address };
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** The preview as an HTML table (column letters, row numbers, the cells with their formats). */
export function sheetHtml(p: SheetPreview, width?: number) {
  const letters = (n: number): string => (n > 26 ? letters(Math.floor((n - 1) / 26)) : "") + String.fromCharCode(65 + ((n - 1) % 26));
  const hidden = new Set<string>();
  const spans = new Map<string, { rs: number; cs: number }>();
  for (const m of p.merges) {
    const [a, b] = m.split(":");
    const s = parseRef(a);
    const e = parseRef(b ?? a);
    spans.set(`${s.r},${s.c}`, { rs: e.r - s.r + 1, cs: e.c - s.c + 1 });
    for (let r = s.r; r <= e.r; r++) for (let c = s.c; c <= e.c; c++) if (r !== s.r || c !== s.c) hidden.add(`${r},${c}`);
  }
  const css = (s?: CellStyle, numeric?: boolean) =>
    [
      s?.b && "font-weight:700",
      s?.i && "font-style:italic",
      s?.u && "text-decoration:underline",
      s?.sz && `font-size:${s.sz}pt`,
      s?.color && `color:${s.color}`,
      s?.fill && `background:${s.fill}`,
      `text-align:${s?.h === "centerContinuous" ? "center" : (s?.h ?? (numeric ? "right" : "left"))}`,
      s?.v && `vertical-align:${s.v === "middle" ? "middle" : s.v}`,
      s?.wrap ? "white-space:pre-wrap" : "white-space:nowrap",
      s?.indent && `padding-left:${4 + s.indent * 9}px`,
      s?.bt && `border-top:${s.bt}`,
      s?.br && `border-right:${s.br}`,
      s?.bb && `border-bottom:${s.bb}`,
      s?.bl && `border-left:${s.bl}`,
    ]
      .filter(Boolean)
      .join(";");
  const head = `<tr><th class="c"></th>${p.widths.map((w, c) => `<th class="h" style="width:${w}px;min-width:${w}px;max-width:${w}px">${letters(c + 1)}</th>`).join("")}</tr>`;
  const body = p.cells
    .map((row, r) => `<tr${p.heights[r] ? ` style="height:${p.heights[r]}px"` : ""}><th class="n">${r + 1}</th>${row
      .map((t, c) => {
        const key = `${r + 1},${c + 1}`;
        if (hidden.has(key)) return "";
        const span = spans.get(key);
        const numeric = !!t && /^[-+(]?[\d\s.,]+%?\)?$|^[-+]?[$€£]/.test(t.trim());
        return `<td${span ? ` rowspan="${span.rs}" colspan="${span.cs}"` : ""} style="${css(p.styles[r]?.[c], numeric)}">${t ? esc(t) : ""}</td>`;
      })
      .join("")}</tr>`)
    .join("");
  return `<!doctype html><html><head><meta charset="utf-8"><style>
body{margin:0;font:11pt Calibri,Carlito,"Segoe UI",Arial,sans-serif;color:#000;background:#fff}
table{border-collapse:collapse;table-layout:fixed${width ? `;min-width:${width}px` : ""}}
td{border:1px solid #e3e3e3;padding:1px 4px;overflow:hidden;text-overflow:clip;height:20px;max-width:0}
th{background:#f3f3f3;color:#555;font-weight:400;font-size:9pt;border:1px solid #d4d4d4;padding:0 4px;position:sticky;top:0;z-index:1}
th.n{position:sticky;left:0;width:36px;min-width:36px;text-align:center;z-index:2}th.c{position:sticky;left:0;top:0;z-index:3;width:36px}
</style></head><body><table>${head}${body}</table></body></html>`;
}
