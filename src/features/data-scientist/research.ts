/**
 * A research (Data Scientist): the semantic models it reads, the analysis (a markdown text) and a small
 * workbook of sheets, as in Excel: a pivot (rows, columns, values, filters: the DAX is written by code),
 * a chart of a pivot's figures, or a calculation (cells with Excel formulas, which reach the pivots with
 * GETPIVOTDATA). Kept in this browser; the queries run again from a sheet's refresh button, and the
 * workbook downloads as Excel with every pivot connected to its model.
 */
import { dax, type Ref, type Rows } from "@/services/pbi";
import type { ModelInfo } from "@/services/model";
import { queryDax, type FilterSpec } from "@/lib/query-spec";

export const CHARTS = ["column", "bar", "line", "area", "pie", "scatter"] as const;
export type ChartKind = (typeof CHARTS)[number];

/** A pivot's parameters: rows and columns (columns of the model), values (measures), filters. */
export type PivotSpec = { rows: string[]; columns: string[]; values: string[]; filters: FilterSpec[] };

type SheetBase = { id: string; name: string; note?: string };
type DataSheet = SheetBase & { datasetId: string; rows?: Rows; truncated?: number; error?: string; at?: number };
/** A pivot table: its parameters, its rows once run. */
export type PivotSheet = DataSheet & { kind: "pivot"; spec: PivotSpec };
/** A chart of a pivot (rows: the categories, columns: the series); older canvases: a free DAX query (dax, x, y). */
export type ChartSheet = DataSheet & { kind: "chart"; chart: ChartKind; spec?: PivotSpec; dax?: string; x?: string; y?: string[] };
/** A calculation: a grid of cells (text, numbers, or "=" formulas). */
export type CalcSheet = SheetBase & { kind: "calc"; cells: string[][] };
export type Sheet = PivotSheet | ChartSheet | CalcSheet;

export type ResearchEntry = { id: string; name: string; datasetIds: string[]; text: string; sheets: Sheet[]; updatedAt: number };

export const newResearch = (name: string): ResearchEntry => ({ id: crypto.randomUUID(), name, datasetIds: [], text: "", sheets: [], updatedAt: Date.now() });

/** A sheet name as Excel takes it: at most 28 characters, none of : \ / ? * [ ] (the "(data)" sheet adds to it). */
export const sheetName = (name: string) => name.replace(/[:\\/?*[\]']/g, " ").replace(/\s+/g, " ").trim().slice(0, 28) || "Sheet";

/**
 * A research as this version reads it: canvases of earlier versions become sheets (a pivot card a pivot
 * sheet, a chart card a chart sheet keeping its DAX), their text blocks the text.
 */
export function readResearch(r: Partial<ResearchEntry> & { id: string; name: string; blocks?: { kind: string; text?: string }[]; cards?: any[] }): ResearchEntry {
  const text = typeof r.text === "string" ? r.text : (r.blocks ?? []).filter((b) => b.kind === "markdown").map((b) => b.text ?? "").join("\n\n");
  const fromCards = (r.cards ?? []).map((c): Sheet => {
    const base = { id: c.id, name: sheetName(c.title ?? "Sheet"), datasetId: c.datasetId, rows: c.rows, truncated: c.truncated, error: c.error, at: c.at };
    return c.kind === "chart" ? { ...base, kind: "chart", chart: c.chart === "column" || !CHARTS.includes(c.chart) ? "column" : c.chart, dax: c.dax, x: c.x, y: c.y } : { ...base, kind: "pivot", spec: c.spec };
  });
  return { id: r.id, name: r.name, datasetIds: Array.isArray(r.datasetIds) ? r.datasetIds : [], text, sheets: Array.isArray(r.sheets) ? r.sheets : fromCards, updatedAt: r.updatedAt ?? Date.now() };
}

/** A field's name in a query result ('T'[col] → col). */
export const resultName = (field: string) => field.match(/\[([^\]]+)\]\s*$/)?.[1] ?? field;

/** The DAX of a pivot: its rows and columns grouped, its values computed, its filters applied. */
export const pivotDax = (spec: PivotSpec, info: ModelInfo | undefined) => queryDax({ values: spec.values, group_by: [...spec.rows, ...spec.columns], filters: spec.filters }, info);

/** The DAX a data sheet runs (a chart of an older canvas: its own query). */
export const sheetDax = (sheet: PivotSheet | ChartSheet, info: ModelInfo | undefined) => (sheet.kind === "chart" && !sheet.spec ? sheet.dax ?? "" : pivotDax(sheet.spec!, info));

/** Runs a data sheet's query: the sheet with its rows (or the error). */
export async function runSheet<S extends PivotSheet | ChartSheet>(sheet: S, model: Ref & { info?: ModelInfo }): Promise<S> {
  try {
    const rows = await dax(model, sheetDax(sheet, model.info));
    return { ...sheet, rows, truncated: rows.truncated, error: undefined, at: Date.now() };
  } catch (e) {
    return { ...sheet, rows: undefined, error: e instanceof Error ? e.message : String(e), at: Date.now() };
  }
}

/** The x and series of a chart: a pivot's first row field and its values (or the columns of an older chart). */
export function chartAxes(sheet: ChartSheet) {
  if (!sheet.spec) return { x: sheet.x ?? "", y: sheet.y ?? [] };
  const p = pivot(sheet.spec, sheet.rows ?? []);
  return { x: p.rowLabels.join(" / ") || "Total", y: p.head.map((h) => h.label || "Value"), data: p.lines.map((l) => ({ [p.rowLabels.join(" / ") || "Total"]: l.labels.join(" / ") || "Total", ...Object.fromEntries(p.head.map((h) => [h.label || "Value", l.cells.get(h.id)])) })) };
}

/**
 * The pivot of a sheet's rows: one line per combination of its row fields, one column per combination
 * of its column fields and value (cells: the value, blank when missing).
 */
export function pivot(spec: PivotSpec, rows: Rows) {
  const rowKeys = spec.rows.map(resultName);
  const colKeys = spec.columns.map(resultName);
  const values = rows.length ? Object.keys(rows[0]).filter((k) => !rowKeys.includes(k) && !colKeys.includes(k)) : [];
  const key = (r: Rows[number], keys: string[]) => keys.map((k) => String(r[k] ?? "")).join(" · ");
  const columns = colKeys.length ? [...new Set(rows.map((r) => key(r, colKeys)))].sort() : [""];
  const lines = new Map<string, { labels: string[]; cells: Map<string, unknown> }>();
  for (const r of rows) {
    const k = key(r, rowKeys);
    const line = lines.get(k) ?? lines.set(k, { labels: rowKeys.map((x) => String(r[x] ?? "")), cells: new Map() }).get(k)!;
    for (const v of values) line.cells.set(`${key(r, colKeys)}|${v}`, r[v]);
  }
  const head = columns.flatMap((c) => values.map((v) => ({ id: `${c}|${v}`, label: [c, values.length > 1 || !c ? v : ""].filter(Boolean).join(" · ") })));
  return { rowLabels: rowKeys, head, lines: [...lines.values()] };
}

/**
 * GETPIVOTDATA on a sheet's rows, as Excel computes it on the PivotTable built from them: the sum of the
 * value over the rows matching every field = item pair (fields by name, items compared as text).
 */
export function pivotData(sheet: PivotSheet | ChartSheet, value: string, pairs: [string, unknown][]) {
  const rows = sheet.rows ?? [];
  if (!rows.length) return undefined;
  const keys = Object.keys(rows[0]);
  const find = (name: string) => keys.find((k) => k.toLowerCase() === resultName(name).toLowerCase().trim());
  const v = find(value);
  if (!v) throw new Error(`#REF! "${value}" is not a value of "${sheet.name}" (${keys.join(", ")})`);
  const filters = pairs.map(([f, item]) => {
    const k = find(String(f));
    if (!k) throw new Error(`#REF! "${f}" is not a field of "${sheet.name}" (${keys.join(", ")})`);
    return [k, String(item ?? "").toLowerCase()] as const;
  });
  const matching = rows.filter((r) => filters.every(([k, item]) => String(r[k] ?? "").toLowerCase() === item));
  if (!matching.length) throw new Error(`#REF! no row of "${sheet.name}" for ${filters.map(([k, i]) => `${k} = ${i}`).join(", ")}`);
  return matching.reduce((n, r) => n + (Number(r[v]) || 0), 0);
}
