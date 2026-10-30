/**
 * A research (Data Scientist): the semantic models it reads, and a canvas — a markdown text (the answer)
 * and cards below it, each ONE chart or ONE pivot table fed by a DAX query on one model. Kept in this
 * browser; the queries run again from a card's refresh button.
 */
import { dax, type Ref, type Rows } from "@/services/pbi";
import type { ModelInfo } from "@/services/model";
import { queryDax, type FilterSpec } from "@/lib/query-spec";

export const CHARTS = ["bar", "column", "line", "area", "pie", "scatter"] as const;
export type ChartKind = (typeof CHARTS)[number];

/** A pivot's parameters: rows and columns (columns of the model), values (measures), filters. */
export type PivotSpec = { rows: string[]; columns: string[]; values: string[]; filters: FilterSpec[] };

type CardBase = { id: string; title: string; datasetId: string; rows?: Rows; truncated?: number; error?: string; at?: number };
/** A chart drawn in the app (not Power BI) from a free DAX query: `x` the category column, `y` the value columns. */
export type ChartCard = CardBase & { kind: "chart"; chart: ChartKind; dax: string; x: string; y: string[] };
/** A pivot table from its parameters (the DAX is written by code). */
export type PivotCard = CardBase & { kind: "pivot"; spec: PivotSpec };
export type Card = ChartCard | PivotCard;

export type ResearchEntry = { id: string; name: string; datasetIds: string[]; text: string; cards: Card[]; updatedAt: number };

export const newResearch = (name: string): ResearchEntry => ({ id: crypto.randomUUID(), name, datasetIds: [], text: "", cards: [], updatedAt: Date.now() });

/** A research as this version reads it (older ones: their text blocks become the text, the rest is dropped). */
export function readResearch(r: Partial<ResearchEntry> & { id: string; name: string; blocks?: { kind: string; text?: string }[] }): ResearchEntry {
  const text = typeof r.text === "string" ? r.text : (r.blocks ?? []).filter((b) => b.kind === "markdown").map((b) => b.text ?? "").join("\n\n");
  return { id: r.id, name: r.name, datasetIds: Array.isArray(r.datasetIds) ? r.datasetIds : [], text, cards: Array.isArray(r.cards) ? r.cards : [], updatedAt: r.updatedAt ?? Date.now() };
}

/** A field's name in a query result ('T'[col] → col). */
export const resultName = (field: string) => field.match(/\[([^\]]+)\]\s*$/)?.[1] ?? field;

/** The DAX of a pivot: its rows and columns grouped, its values computed, its filters applied. */
export const pivotDax = (spec: PivotSpec, info: ModelInfo | undefined) => queryDax({ values: spec.values, group_by: [...spec.rows, ...spec.columns], filters: spec.filters }, info);

/** The DAX a card runs. */
export const cardDax = (card: Card, info: ModelInfo | undefined) => (card.kind === "chart" ? card.dax : pivotDax(card.spec, info));

/** Runs a card's query: the card with its rows (or the error). */
export async function runCard<C extends Card>(card: C, model: Ref & { info?: ModelInfo }): Promise<C> {
  try {
    const rows = await dax(model, cardDax(card, model.info));
    if (card.kind === "chart" && rows.length) {
      const missing = [card.x, ...card.y].filter((c) => !(c in rows[0]));
      if (missing.length) throw new Error(`The query returns no column ${missing.map((m) => `"${m}"`).join(", ")} (its columns: ${Object.keys(rows[0]).join(", ")}).`);
    }
    return { ...card, rows, truncated: rows.truncated, error: undefined, at: Date.now() };
  } catch (e) {
    return { ...card, rows: undefined, error: e instanceof Error ? e.message : String(e), at: Date.now() };
  }
}

/**
 * The pivot of a card's rows: one line per combination of its row fields, one column per combination
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
