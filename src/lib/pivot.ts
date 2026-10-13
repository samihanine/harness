/**
 * TCD builder: a pivot table described as data (rows, columns, values, filters), turned into one DAX
 * query by `toDax` — deterministic, no AI involved. Shared by the screen and its AI tools.
 */
import type { ModelInfo } from "./model";

export type Col = { table: string; column: string };
export const AGGREGATIONS = { sum: "Sum", count: "Count", distinct: "Distinct count", average: "Average", min: "Min", max: "Max" } as const;
export type Aggregation = keyof typeof AGGREGATIONS;
/** A value: a measure, or a column with an aggregation. */
export type Value = { table: string; measure?: string; column?: string; aggregation?: Aggregation };

/** Filter operators: on the values of the column, or a condition (text, number, date). */
export const OPERATORS = {
  in: "is one of",
  notIn: "is not one of",
  contains: "contains",
  notContains: "does not contain",
  startsWith: "starts with",
  eq: "=",
  ne: "≠",
  gt: ">",
  ge: "≥",
  lt: "<",
  le: "≤",
  between: "between",
  blank: "is empty",
  notBlank: "is not empty",
} as const;
export type Operator = keyof typeof OPERATORS;
export type ColumnKind = "text" | "number" | "date";
/** Operators offered for a kind of column. */
export const OPERATORS_OF: Record<ColumnKind, Operator[]> = {
  text: ["in", "notIn", "contains", "notContains", "startsWith", "eq", "ne", "blank", "notBlank"],
  number: ["in", "notIn", "eq", "ne", "gt", "ge", "lt", "le", "between", "blank", "notBlank"],
  date: ["in", "notIn", "eq", "ne", "gt", "ge", "lt", "le", "between", "blank", "notBlank"],
};
/** `values` for in / notIn; `value` (and `value2` for between) for the conditions. Older filters have no `op` ("in"). */
export type Filter = Col & { op?: Operator; values: unknown[]; value?: string; value2?: string };
export type Spec = { rows: Col[]; columns: Col[]; values: Value[]; filters: Filter[] };
export const EMPTY: Spec = { rows: [], columns: [], values: [], filters: [] };

export const ref = (c: Col) => `'${c.table.replace(/'/g, "''")}'[${c.column.replace(/]/g, "]]")}]`;
const text = (v: unknown) => `"${String(v).replace(/"/g, '""')}"`;
const literal = (v: unknown) => (typeof v === "number" ? String(v) : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : text(v));
const DAX_AGG: Record<Aggregation, string> = { sum: "SUM", count: "COUNTA", distinct: "DISTINCTCOUNT", average: "AVERAGE", min: "MIN", max: "MAX" };
export const labelOf = (v: Value) => (v.measure ? v.measure : `${AGGREGATIONS[v.aggregation ?? "sum"]} of ${v.column}`);
const exprOf = (v: Value) => (v.measure ? `'${v.table.replace(/'/g, "''")}'[${v.measure.replace(/]/g, "]]")}]` : `${DAX_AGG[v.aggregation ?? "sum"]}(${ref({ table: v.table, column: v.column! })})`);

/** Text, number or date column, from the model's types. */
export function kindOf(info: ModelInfo | undefined, c: Col): ColumnKind {
  const type = info?.tables.find((t) => t.name === c.table)?.columns.find((x) => x.name === c.column)?.type ?? "";
  return /date|time/i.test(type) ? "date" : /int|double|decimal|currency|number/i.test(type) ? "number" : "text";
}

/** A typed condition value: number, DATE(y, m, d), or text. */
function typed(value: string | undefined, kind: ColumnKind) {
  const v = (value ?? "").trim();
  if (kind === "number" && v !== "" && !Number.isNaN(Number(v.replace(",", ".")))) return String(Number(v.replace(",", ".")));
  const d = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (kind === "date" && d) return `DATE(${Number(d[1])}, ${Number(d[2])}, ${Number(d[3])})`;
  return text(v);
}

/** Is the filter complete enough to apply (an empty value list or condition does not filter)? */
export const active = (f: Filter) => {
  const op = f.op ?? "in";
  if (op === "in" || op === "notIn") return f.values.length > 0;
  if (op === "blank" || op === "notBlank") return true;
  if (op === "between") return !!f.value?.trim() && !!f.value2?.trim();
  return !!f.value?.trim();
};

/** One filter as a SUMMARIZECOLUMNS filter table. */
function filterDax(f: Filter, kind: ColumnKind) {
  const c = ref(f);
  const op = f.op ?? "in";
  if (op === "in") return `TREATAS({${f.values.map(literal).join(", ")}}, ${c})`;
  const condition = {
    notIn: () => `NOT (${c} IN {${f.values.map(literal).join(", ")}})`,
    contains: () => `CONTAINSSTRING(${c}, ${text(f.value)})`,
    notContains: () => `NOT CONTAINSSTRING(${c}, ${text(f.value)})`,
    startsWith: () => `LEFT(${c}, ${(f.value ?? "").length}) = ${text(f.value)}`,
    eq: () => `${c} = ${typed(f.value, kind)}`,
    ne: () => `${c} <> ${typed(f.value, kind)}`,
    gt: () => `${c} > ${typed(f.value, kind)}`,
    ge: () => `${c} >= ${typed(f.value, kind)}`,
    lt: () => `${c} < ${typed(f.value, kind)}`,
    le: () => `${c} <= ${typed(f.value, kind)}`,
    between: () => `${c} >= ${typed(f.value, kind)} && ${c} <= ${typed(f.value2, kind)}`,
    blank: () => `ISBLANK(${c})`,
    notBlank: () => `NOT ISBLANK(${c})`,
  }[op]();
  return `FILTER(ALL(${c}), ${condition})`;
}

/** Short text of a filter (chips, AI state). */
export function filterText(f: Filter) {
  const op = f.op ?? "in";
  if (!active(f)) return `${f.column} (all)`;
  if (op === "in" || op === "notIn") return `${f.column} ${op === "in" ? "=" : "≠"} ${f.values.slice(0, 3).join(", ")}${f.values.length > 3 ? "…" : ""}`;
  if (op === "blank" || op === "notBlank") return `${f.column} ${OPERATORS[op]}`;
  return `${f.column} ${OPERATORS[op]} ${f.value}${op === "between" ? ` and ${f.value2}` : ""}`;
}

/** The pivot as one SUMMARIZECOLUMNS: rows and columns grouped, filters as filter tables, values as named expressions. */
export function toDax(spec: Spec, info?: ModelInfo) {
  const groups = [...spec.rows, ...spec.columns].map(ref);
  const filters = spec.filters.filter(active).map((f) => filterDax(f, kindOf(info, f)));
  const values = spec.values.map((v) => `"${labelOf(v).replace(/"/g, '""')}", ${exprOf(v)}`);
  const args = [...groups, ...filters, ...values];
  if (!groups.length && !values.length) return "";
  return `EVALUATE\nSUMMARIZECOLUMNS(\n    ${args.join(",\n    ")}\n)${groups.length ? `\nORDER BY ${groups.join(", ")}` : ""}`;
}

/** The pivot in a few lines (AI state). */
export const specText = (spec: Spec) =>
  [
    `rows: ${spec.rows.map(ref).join(", ") || "none"}`,
    `columns: ${spec.columns.map(ref).join(", ") || "none"}`,
    `values: ${spec.values.map((v) => (v.measure ? `'${v.table}'[${v.measure}]` : `${AGGREGATIONS[v.aggregation ?? "sum"]}(${ref({ table: v.table, column: v.column! })})`)).join(", ") || "none"}`,
    `filters: ${spec.filters.map((f) => `'${f.table}'[${f.column}] ${filterText(f)}`).join("; ") || "none"}`,
  ].join("\n");
