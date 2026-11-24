/**
 * A Data Scientist analysis: its sources (DAX queries on semantic models, sheets of uploaded Excel / CSV
 * files) loaded as SQL tables, its steps (SQL views combining them) and what it shows (blocks: text, key
 * figures, tables, charts, pivots), each block keeping the rows it showed. Rebuilt on the server from this
 * document (sources reloaded, views created again), kept by the browser.
 */

export type Source =
  | { name: string; kind: "dax"; datasetId: string; groupId?: string; datasetName: string; query: string; rows?: number }
  | { name: string; kind: "file"; fileId: string; fileName: string; sheet?: string; rows?: number };

export type Step = { name: string; sql: string; description?: string };

export type Rows = Record<string, unknown>[];

export type ChartKind = "bar" | "column" | "line" | "area" | "pie" | "scatter";

export type AnalysisBlock =
  | { id: string; type: "text"; markdown: string }
  | { id: string; type: "kpis"; title?: string; sql: string; items: { label: string; column: string; format?: string }[]; rows?: Rows }
  | { id: string; type: "table"; title?: string; sql: string; rows?: Rows; columns?: string[]; truncated?: boolean }
  | { id: string; type: "chart"; title?: string; sql: string; chart: { kind: ChartKind; x: string; y: string[]; stacked?: boolean }; rows?: Rows }
  | { id: string; type: "pivot"; title?: string; sql: string; pivot: { rows: string[]; columns?: string; value: string; aggregate: "sum" | "avg" | "count" | "min" | "max" }; rows?: Rows };

export type Analysis = { title: string; sources: Source[]; steps: Step[]; blocks: AnalysisBlock[]; updatedAt?: string };

export const emptyAnalysis = (title = "New analysis"): Analysis => ({ title, sources: [], steps: [], blocks: [] });

/** A pivot of rows: one line per `rows` combination, one column per value of `columns`, the value aggregated. */
export function pivotRows(rows: Rows, p: Extract<AnalysisBlock, { type: "pivot" }>["pivot"]) {
  const key = (r: Record<string, unknown>) => p.rows.map((c) => String(r[c] ?? "")).join("\u0001");
  const cols = p.columns ? [...new Set(rows.map((r) => String(r[p.columns!] ?? "")))].sort() : ["Value"];
  const groups = new Map<string, { head: Record<string, unknown>; cells: Map<string, number[]> }>();
  for (const r of rows) {
    const k = key(r);
    if (!groups.has(k)) groups.set(k, { head: Object.fromEntries(p.rows.map((c) => [c, r[c]])), cells: new Map() });
    const col = p.columns ? String(r[p.columns] ?? "") : "Value";
    const g = groups.get(k)!;
    if (!g.cells.has(col)) g.cells.set(col, []);
    g.cells.get(col)!.push(Number(r[p.value]) || 0);
  }
  const agg = (v: number[] | undefined) => {
    if (!v?.length) return null;
    switch (p.aggregate) {
      case "avg":
        return v.reduce((a, b) => a + b, 0) / v.length;
      case "count":
        return v.length;
      case "min":
        return Math.min(...v);
      case "max":
        return Math.max(...v);
      default:
        return v.reduce((a, b) => a + b, 0);
    }
  };
  return { columns: cols, lines: [...groups.values()].map((g) => ({ ...g.head, ...Object.fromEntries(cols.map((c) => [c, agg(g.cells.get(c))])) })) };
}
