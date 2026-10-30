/**
 * A data question as a short spec — values, group by, filters, top N, order — written as DAX by this
 * code (SUMMARIZECOLUMNS, TREATAS, TOPN), names checked against the model first: the AI fills the
 * spec, never the syntax. What the spec cannot say (time intelligence, a filter on a measure, a
 * calculation) is left to raw DAX (run_dax).
 */
import { resolveField, type ModelInfo } from "@/services/model";

export type FilterSpec = {
  field: string;
  /** in (default, with values), not_in, >, >=, <, <=, between (value, to), contains (value). */
  op?: "in" | "not_in" | ">" | ">=" | "<" | "<=" | "between" | "contains";
  values?: (string | number | boolean)[];
  value?: string | number;
  to?: string | number;
};
export type QuerySpec = {
  values: string[];
  group_by?: string[];
  filters?: FilterSpec[];
  top?: { count: number; by?: string; ascending?: boolean };
  order_by?: { field: string; descending?: boolean }[];
};
/** Measures of the report (not of the model), defined in the query. */
export type DefinedMeasure = { table: string; name: string; expression?: string };

const AGGREGATIONS: Record<string, string> = { sum: "SUM", avg: "AVERAGE", average: "AVERAGE", count: "COUNT", countnonnull: "COUNTA", distinctcount: "DISTINCTCOUNT", min: "MIN", max: "MAX" };
const norm = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();
const ref = (table: string, name: string) => `'${table.replace(/'/g, "''")}'[${name}]`;

/** A value or a measure: [Measure] (model or report), Sum('T'[col]) / Avg / Count / DistinctCount / Min / Max. */
function valueOf(spec: string, info: ModelInfo | undefined, defined: DefinedMeasure[]) {
  const m = spec.trim().match(/^(\w+)\((.+)\)$/);
  if (m) {
    const fn = AGGREGATIONS[m[1].toLowerCase()];
    if (!fn) throw new Error(`"${m[1]}" is not an aggregation (${Object.keys(AGGREGATIONS).join(", ")}).`);
    const c = resolveField(info, m[2]);
    if (c.kind === "measure") throw new Error(`[${c.name}] is a measure: use it without ${m[1]}().`);
    return { label: `${m[1][0].toUpperCase()}${m[1].slice(1).toLowerCase()} of ${c.name}`, expr: `${fn}(${ref(c.table, c.name)})` };
  }
  const bare = spec.trim().match(/\[([^\]]+)\]$/)?.[1] ?? spec.trim();
  const own = defined.find((x) => norm(x.name) === norm(bare));
  if (own) return { label: own.name, expr: `[${own.name}]` };
  const f = resolveField(info, spec);
  if (f.kind === "column") throw new Error(`${ref(f.table, f.name)} is a column: as a value write Sum(${ref(f.table, f.name)}) (or Count, Avg, Min, Max), or put it in group_by.`);
  return { label: f.name, expr: `[${f.name}]` };
}

function columnOf(spec: string, info: ModelInfo | undefined) {
  const f = resolveField(info, spec);
  if (f.kind === "measure") throw new Error(`[${f.name}] is a measure: it goes in values (a filter on a measure needs run_dax).`);
  const type = info?.tables.find((t) => t.name === f.table)?.columns.find((c) => c.name === f.name)?.type;
  return { ref: ref(f.table, f.name), type: type?.toLowerCase() };
}

/** A literal of the column's type (a date as DATE(y, m, d), a number as is, text in double quotes). */
export function literal(v: unknown, type?: string) {
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (type?.startsWith("date")) {
    const d = String(v).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!d) throw new Error(`"${v}" is not a date: write it YYYY-MM-DD.`);
    return `DATE(${+d[1]}, ${+d[2]}, ${+d[3]})`;
  }
  const numeric = type && /int|double|decimal|number|currency/.test(type);
  if (typeof v === "number" || (numeric && /^-?\d+(\.\d+)?$/.test(String(v)))) return String(v);
  return `"${String(v).replace(/"/g, '""')}"`;
}

function filterOf(f: FilterSpec, info: ModelInfo | undefined) {
  const c = columnOf(f.field, info);
  const op = f.op ?? "in";
  const list = () => {
    const values = f.values ?? (f.value !== undefined ? [f.value] : []);
    if (!values.length) throw new Error(`Filter on ${c.ref}: give "values".`);
    return `{${values.map((v) => literal(v, c.type)).join(", ")}}`;
  };
  const one = (v: unknown, what = "value") => {
    if (v === undefined || v === "") throw new Error(`Filter on ${c.ref} (${op}): give "${what}".`);
    return literal(v, c.type);
  };
  const all = `ALL(${c.ref})`;
  switch (op) {
    case "in":
      return `TREATAS(${list()}, ${c.ref})`;
    case "not_in":
      return `FILTER(${all}, NOT ${c.ref} IN ${list()})`;
    case "between":
      return `FILTER(${all}, ${c.ref} >= ${one(f.value)} && ${c.ref} <= ${one(f.to, "to")})`;
    case "contains":
      return `FILTER(${all}, CONTAINSSTRING(${c.ref}, ${literal(String(f.value ?? ""), "string")}))`;
    default:
      return `FILTER(${all}, ${c.ref} ${op} ${one(f.value ?? f.values?.[0])})`;
  }
}

/** The DAX query of a spec (report measures defined first, all of them: they may use each other). */
export function queryDax(spec: QuerySpec, info: ModelInfo | undefined, defined: DefinedMeasure[] = []) {
  if (!spec.values?.length && !spec.group_by?.length) throw new Error('Give "values" (measures) and/or "group_by" (columns).');
  const values = (spec.values ?? []).map((v) => valueOf(v, info, defined));
  const groups = (spec.group_by ?? []).map((g) => columnOf(g, info).ref);
  const filters = (spec.filters ?? []).map((f) => filterOf(f, info));
  const table = `SUMMARIZECOLUMNS(${[...groups, ...filters, ...values.map((v) => `"${v.label.replace(/"/g, '""')}", ${v.expr}`)].join(", ")})`;
  /** A result column: a value's label, or a grouped column. */
  const resultColumn = (field: string) => {
    const v = values.find((x) => norm(x.label) === norm(field.replace(/^\[|\]$/g, "")));
    if (v) return `[${v.label}]`;
    try {
      const c = columnOf(field, info).ref;
      if (groups.includes(c)) return c;
    } catch {}
    const v2 = (() => {
      try {
        return valueOf(field, info, defined);
      } catch {
        return undefined;
      }
    })();
    const same = v2 && values.find((x) => x.expr === v2.expr);
    if (same) return `[${same.label}]`;
    throw new Error(`"${field}" is not in the result: order by one of ${[...values.map((x) => `[${x.label}]`), ...groups].join(", ")}.`);
  };
  let body = table;
  let order = (spec.order_by ?? []).map((o) => `${resultColumn(o.field)} ${o.descending ? "DESC" : "ASC"}`);
  if (spec.top) {
    if (!groups.length) throw new Error('"top" needs group_by (the items to rank).');
    if (!values.length && !spec.top.by) throw new Error('"top" needs a value to rank by.');
    const by = spec.top.by ? resultColumn(spec.top.by) : `[${values[0].label}]`;
    const dir = spec.top.ascending ? "ASC" : "DESC";
    body = `TOPN(${Math.max(1, Math.round(spec.top.count))}, ${table}, ${by}, ${dir})`;
    if (!order.length) order = [`${by} ${dir}`];
  }
  const defs = defined.filter((m) => m.expression).map((m) => `MEASURE ${ref(m.table, m.name)} = ${m.expression}`);
  return [defs.length && `DEFINE\n${defs.join("\n")}`, `EVALUATE\n${body}`, order.length && `ORDER BY ${order.join(", ")}`].filter(Boolean).join("\n");
}

export { QUERY_HELP } from "@/prompts/dax";
