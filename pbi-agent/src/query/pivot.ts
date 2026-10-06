/**
 * Pivot tables over one or several sources. Each source is queried on its own (DAX or SQL,
 * generated here), then results are joined on the row/column labels. Computed values combine
 * values of different sources (e.g. the gap between two systems): `{Sales A} - {Sales B}`.
 */
import { z } from "zod";
import type { QueryResult } from "@/pbi/api";
import type { Source } from "./engine";
import { runQuery } from "./engine";

export const AGGREGATES = ["sum", "avg", "count", "distinctcount", "min", "max"] as const;

const field = z.string().describe("'Table'[Column] (or [Measure] for values)");

export const pivotSchema = z.object({
  /** Row / column labels; each is bound to a column of every source it applies to. */
  rows: z.array(z.object({ label: z.string(), bindings: z.record(z.string(), field).describe("source id → 'Table'[Column]") })).default([]),
  columns: z.array(z.object({ label: z.string(), bindings: z.record(z.string(), field) })).default([]),
  values: z
    .array(
      z.object({
        label: z.string(),
        source: z.string().optional().describe("source id (not for computed values)"),
        /** A model measure "[Name]", or a column aggregated with `aggregate`. */
        field: field.optional(),
        aggregate: z.enum(AGGREGATES).optional(),
        /** Native expression of the source: DAX for remote, SQL aggregate for local. */
        expression: z.string().optional(),
        /** Combination of other values by label: `{Sales A} - {Sales B}`, DIVIDE(a, b), ABS(x). */
        computed: z.string().optional(),
        format: z.enum(["number", "integer", "percent", "currency"]).optional(),
      }),
    )
    .default([]),
  filters: z.array(z.object({ source: z.string(), field, values: z.array(z.union([z.string(), z.number(), z.boolean()])) })).default([]),
});

export type Pivot = z.infer<typeof pivotSchema>;
export const emptyPivot = (): Pivot => ({ rows: [], columns: [], values: [], filters: [] });

/** 'Table'[Column] → { table, column }. */
export function parseField(ref: string) {
  const match = /^'?(.*?)'?\[(.+)\]$/.exec(ref.trim());
  return match ? { table: match[1], column: match[2] } : { table: "", column: ref.trim() };
}
const daxField = (ref: string) => {
  const { table, column } = parseField(ref);
  return table ? `'${table.replace(/'/g, "''")}'[${column}]` : `[${column}]`;
};
const sqlField = (ref: string) => {
  const { table, column } = parseField(ref);
  return table ? `[${table}].[${column}]` : `[${column}]`;
};
const literal = (v: string | number | boolean, sql: boolean) =>
  typeof v === "string" ? (sql ? `'${v.replace(/'/g, "''")}'` : `"${v.replace(/"/g, '""')}"`) : String(v).toUpperCase();

const DAX_AGG: Record<(typeof AGGREGATES)[number], string> = {
  sum: "SUM",
  avg: "AVERAGE",
  count: "COUNT",
  distinctcount: "DISTINCTCOUNT",
  min: "MIN",
  max: "MAX",
};
const SQL_AGG: Record<(typeof AGGREGATES)[number], string> = {
  sum: "SUM",
  avg: "AVG",
  count: "COUNT",
  distinctcount: "COUNT(DISTINCT",
  min: "MIN",
  max: "MAX",
};

type Value = Pivot["values"][number];

function valueExpr(value: Value, source: Source) {
  if (value.expression) return value.expression;
  if (!value.field) throw new Error(`Value "${value.label}" has no field.`);
  const { table, column } = parseField(value.field);
  if (!table) {
    // Model measure: DAX reference, or the stored SQL expression of a local measure.
    if (source.language === "DAX") return `[${column}]`;
    const measure = source.model.tables.flatMap((t) => t.measures).find((m) => m.name === column);
    if (!measure?.expression) throw new Error(`Unknown measure [${column}] in ${source.title}.`);
    return measure.expression;
  }
  const agg = value.aggregate ?? "sum";
  if (source.language === "DAX") return `${DAX_AGG[agg]}(${daxField(value.field)})`;
  return agg === "distinctcount" ? `COUNT(DISTINCT ${sqlField(value.field)})` : `${SQL_AGG[agg]}(${sqlField(value.field)})`;
}

/** The query of one source: grouped by its bound dimensions, with its values and filters. */
export function sourceQuery(pivot: Pivot, source: Source) {
  const dims = [...pivot.rows, ...pivot.columns].map((d, i) => {
    const bound = d.bindings[source.id];
    if (!bound) throw new Error(`"${d.label}" is not mapped to a column of ${source.title}.`);
    return { key: `d${i}`, field: bound };
  });
  const values = pivot.values.flatMap((v, i) => (v.source === source.id && !v.computed ? [{ key: `v${i}`, expr: valueExpr(v, source) }] : []));
  const filters = pivot.filters.filter((f) => f.source === source.id && f.values.length);

  if (source.language === "DAX") {
    const args = [
      ...dims.map((d) => daxField(d.field)),
      ...filters.map((f) => `TREATAS({${f.values.map((v) => literal(v, false)).join(", ")}}, ${daxField(f.field)})`),
      ...values.map((v) => `"${v.key}", ${v.expr}`),
    ];
    const select = [...dims.map((d) => `"${d.key}", ${daxField(d.field)}`), ...values.map((v) => `"${v.key}", [${v.key}]`)];
    // SELECTCOLUMNS gives stable result names whatever the model names are.
    return `EVALUATE\nSELECTCOLUMNS(\n  SUMMARIZECOLUMNS(\n    ${args.join(",\n    ")}\n  ),\n  ${select.join(",\n  ")}\n)`;
  }

  // SQL: fields of several tables are joined along the model relationships.
  const tables = [...new Set([...dims.map((d) => d.field), ...filters.map((f) => f.field), ...pivot.values.filter((v) => v.source === source.id && v.field).map((v) => v.field!)].map((f) => parseField(f).table).filter(Boolean))];
  const from = tables[0] ?? source.model.tables[0]?.name;
  const joins = tables.slice(1).map((table) => {
    const rel = source.model.relationships.find(
      (r) => (parseField(r.from).table === from && parseField(r.to).table === table) || (parseField(r.to).table === from && parseField(r.from).table === table),
    );
    if (!rel) throw new Error(`No relationship between ${from} and ${table} in ${source.title}.`);
    return `LEFT JOIN [${table}] ON ${sqlField(rel.from)} = ${sqlField(rel.to)}`;
  });
  const where = filters.map((f) => `${sqlField(f.field)} IN (${f.values.map((v) => literal(v, true)).join(", ")})`);
  return [
    `SELECT ${[...dims.map((d) => `${sqlField(d.field)} AS [${d.key}]`), ...values.map((v) => `${v.expr} AS [${v.key}]`)].join(", ") || "COUNT(*) AS [n]"}`,
    `FROM [${from}]`,
    ...joins,
    where.length ? `WHERE ${where.join(" AND ")}` : "",
    dims.length ? `GROUP BY ${dims.map((d) => sqlField(d.field)).join(", ")}` : "",
  ]
    .filter(Boolean)
    .join("\n");
}

export type PivotResult = {
  /** One row per label combination: d0…dn labels, v0…vn values. */
  rows: Record<string, unknown>[];
  queries: { source: string; query: string; error?: string }[];
};

/** Runs every source, joins on the labels, then evaluates computed values. */
export async function runPivot(pivot: Pivot, sources: Source[]): Promise<PivotResult> {
  const used = sources.filter((s) => pivot.values.some((v) => v.source === s.id && !v.computed));
  const dimCount = pivot.rows.length + pivot.columns.length;
  const merged = new Map<string, Record<string, unknown>>();
  const queries: PivotResult["queries"] = [];
  await Promise.all(
    used.map(async (source) => {
      let query = "";
      try {
        query = sourceQuery(pivot, source);
        const result: QueryResult = await runQuery(source, query);
        queries.push({ source: source.id, query });
        for (const row of result.rows) {
          const labels = Array.from({ length: dimCount }, (_, i) => row[`d${i}`] ?? null);
          const key = JSON.stringify(labels);
          const target = merged.get(key) ?? Object.fromEntries(labels.map((l, i) => [`d${i}`, l]));
          for (const [k, v] of Object.entries(row)) if (k.startsWith("v")) target[k] = v;
          merged.set(key, target);
        }
      } catch (error) {
        queries.push({ source: source.id, query, error: error instanceof Error ? error.message : String(error) });
      }
    }),
  );
  const rows = [...merged.values()];
  pivot.values.forEach((value, i) => {
    if (!value.computed) return;
    const evaluate = compile(value.computed, pivot);
    for (const row of rows) row[`v${i}`] = evaluate(row);
  });
  rows.sort((a, b) => {
    for (let i = 0; i < dimCount; i++) {
      const c = String(a[`d${i}`] ?? "").localeCompare(String(b[`d${i}`] ?? ""), undefined, { numeric: true });
      if (c) return c;
    }
    return 0;
  });
  return { rows, queries };
}

/* ------------------------ computed values (no eval) ----------------------- */

type Node = (row: Record<string, unknown>) => number | null;

/** Parses `{A} - {B}`, `DIVIDE({A}, {B})`, `ABS(…)`, numbers, + - * / and parentheses. */
export function compile(text: string, pivot: Pivot): Node {
  const tokens = text.match(/\{[^}]+\}|\d+(?:\.\d+)?|[A-Za-z_]+|[-+*/(),]/g) ?? [];
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  const num = (v: unknown) => (typeof v === "number" ? v : v === null || v === undefined || v === "" ? null : Number(v));

  const primary = (): Node => {
    const token = next();
    if (!token) throw new Error(`Unexpected end in "${text}"`);
    if (token === "(") {
      const node = expr();
      if (next() !== ")") throw new Error(`Missing ) in "${text}"`);
      return node;
    }
    if (token === "-") {
      const node = primary();
      return (row) => {
        const v = node(row);
        return v === null ? null : -v;
      };
    }
    if (token.startsWith("{")) {
      const label = token.slice(1, -1).trim();
      const index = pivot.values.findIndex((v) => v.label === label);
      if (index < 0) throw new Error(`Unknown value {${label}} (values: ${pivot.values.map((v) => v.label).join(", ")})`);
      return (row) => num(row[`v${index}`]);
    }
    if (/^\d/.test(token)) return () => Number(token);
    const fn = token.toUpperCase();
    if (next() !== "(") throw new Error(`Expected ( after ${token}`);
    const args: Node[] = [];
    if (peek() !== ")")
      do args.push(expr());
      while (peek() === "," && next());
    if (next() !== ")") throw new Error(`Missing ) after ${fn} arguments`);
    if (fn === "ABS") return (row) => ((v) => (v === null ? null : Math.abs(v)))(args[0](row));
    if (fn === "DIVIDE")
      return (row) => {
        const [a, b] = [args[0](row), args[1](row)];
        return a === null || !b ? (args[2] ? args[2](row) : null) : a / b;
      };
    if (fn === "COALESCE") return (row) => args.map((a) => a(row)).find((v) => v !== null) ?? null;
    throw new Error(`Unknown function ${fn} (use ABS, DIVIDE, COALESCE)`);
  };
  const term = (): Node => {
    let left = primary();
    while (peek() === "*" || peek() === "/") {
      const op = next();
      const [l, r] = [left, primary()];
      left = (row) => {
        const [a, b] = [l(row), r(row)];
        return a === null || b === null ? null : op === "*" ? a * b : b === 0 ? null : a / b;
      };
    }
    return left;
  };
  const expr = (): Node => {
    let left = term();
    while (peek() === "+" || peek() === "-") {
      const op = next();
      const [l, r] = [left, term()];
      // Missing on one side counts as 0 (a value absent from one source is a gap).
      left = (row) => {
        const [a, b] = [l(row), r(row)];
        if (a === null && b === null) return null;
        return op === "+" ? (a ?? 0) + (b ?? 0) : (a ?? 0) - (b ?? 0);
      };
    }
    return left;
  };
  const node = expr();
  if (pos < tokens.length) throw new Error(`Unexpected "${tokens[pos]}" in "${text}"`);
  return node;
}

/** Display grid: row labels × (column labels × values). */
export function layout(pivot: Pivot, result: PivotResult) {
  const r = pivot.rows.length;
  const c = pivot.columns.length;
  const colKeys = c ? [...new Set(result.rows.map((row) => JSON.stringify(pivot.columns.map((_, i) => row[`d${r + i}`]))))] : [JSON.stringify([])];
  const grid = new Map<string, Record<string, unknown>>();
  for (const row of result.rows) {
    const rowKey = JSON.stringify(pivot.rows.map((_, i) => row[`d${i}`]));
    const colKey = JSON.stringify(pivot.columns.map((_, i) => row[`d${r + i}`]));
    const line = grid.get(rowKey) ?? {};
    pivot.values.forEach((_, v) => (line[`${colKey}|${v}`] = row[`v${v}`]));
    grid.set(rowKey, line);
  }
  return {
    colKeys: colKeys.map((k) => JSON.parse(k) as unknown[]),
    rows: [...grid.entries()].map(([key, cells]) => ({ labels: JSON.parse(key) as unknown[], cells })),
    cell: (cells: Record<string, unknown>, colKey: unknown[], value: number) => cells[`${JSON.stringify(colKey)}|${value}`],
  };
}
