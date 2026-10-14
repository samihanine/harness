/**
 * Views of an Excel table in the dataset builder: search, filters (field, operator, value), sorts,
 * grouping, hidden fields, column widths — applied in the browser to the rows read from the file.
 */
import { listOf, type Field, type FieldType, type Row } from "./schema";

export type Op = "contains" | "is" | "isNot" | "empty" | "notEmpty" | "gt" | "lt" | "before" | "after" | "true" | "false";
export const OPS: Record<Op, string> = { contains: "contains", is: "is", isNot: "is not", empty: "is empty", notEmpty: "is not empty", gt: ">", lt: "<", before: "before", after: "after", true: "is checked", false: "is not checked" };
export const OPS_OF: Record<FieldType, Op[]> = {
  string: ["contains", "is", "empty", "notEmpty"],
  text: ["contains", "empty", "notEmpty"],
  image: ["empty", "notEmpty"],
  number: ["is", "gt", "lt", "empty", "notEmpty"],
  integer: ["is", "gt", "lt", "empty", "notEmpty"],
  date: ["is", "before", "after", "empty", "notEmpty"],
  option: ["is", "isNot", "empty", "notEmpty"],
  boolean: ["true", "false"],
};
export const NO_VALUE: Op[] = ["empty", "notEmpty", "true", "false"];

/** `values` for option is / is not (any of them), `value` for the other operators. */
export type Filter = { field: string; op: Op; value?: string; values?: string[] };
export type Sort = { field: string; desc?: boolean };
/**
 * `hidden` undefined = the default columns (the first five fields); `group` = an option field;
 * `widths` = column widths (px) set by the user.
 */
export type TableView = { search?: string; filters: Filter[]; sorts: Sort[]; hidden?: string[]; group?: string; widths?: Record<string, number> };
export const EMPTY_VIEW: TableView = { filters: [], sorts: [] };

/** A stored view, from this version or an older one (other shapes): only what can be read is kept. */
export function readView(v: unknown): TableView {
  const o = (v && typeof v === "object" ? v : {}) as Record<string, unknown>;
  return {
    search: typeof o.search === "string" ? o.search : undefined,
    filters: Array.isArray(o.filters) ? (o.filters as Filter[]).filter((f) => f && typeof f.field === "string" && typeof f.op === "string") : [],
    sorts: Array.isArray(o.sorts) ? (o.sorts as Sort[]) : [],
    hidden: Array.isArray(o.hidden) ? (o.hidden as string[]) : undefined,
    group: typeof o.group === "string" ? o.group : undefined,
    widths: o.widths && typeof o.widths === "object" ? (o.widths as Record<string, number>) : undefined,
  };
}

/** Fields shown by a view (the first five when none were chosen). */
export const shownFields = (fields: Field[], view: TableView) => (view.hidden ? fields.filter((f) => !view.hidden!.includes(f.name)) : fields.slice(0, 5));

const empty = (v: unknown) => v === undefined || v === null || v === "" || (Array.isArray(v) && !v.length);

export function matches(row: Row, f: Filter, fields: Field[]) {
  const field = fields.find((x) => x.name === f.field);
  if (!field) return true;
  const v = row[f.field];
  switch (f.op) {
    case "empty":
      return empty(v);
    case "notEmpty":
      return !empty(v);
    case "true":
      return v === true;
    case "false":
      return v !== true;
    case "contains":
      return !f.value || String(v ?? "").toLowerCase().includes(f.value.toLowerCase());
    case "is":
      if (field.type === "option") return !f.values?.length || listOf(v).some((x) => f.values!.includes(x));
      if (field.type === "number" || field.type === "integer") return !f.value || Number(v) === Number(f.value);
      return !f.value || String(v ?? "").toLowerCase() === f.value.toLowerCase();
    case "isNot":
      return !f.values?.length || !listOf(v).some((x) => f.values!.includes(x));
    case "gt":
      return !f.value || (!empty(v) && Number(v) > Number(f.value));
    case "lt":
      return !f.value || (!empty(v) && Number(v) < Number(f.value));
    case "before":
      return !f.value || (!empty(v) && String(v) < f.value);
    case "after":
      return !f.value || (!empty(v) && String(v) > f.value);
  }
}

/** Option values sort in the options' order (Low < Medium < High), empty values last. */
function compare(a: unknown, b: unknown, field?: Field) {
  if (empty(a)) return empty(b) ? 0 : 1;
  if (empty(b)) return -1;
  if (field?.type === "option") {
    const i = (v: unknown) => field.options?.findIndex((o) => o.value === listOf(v)[0]) ?? 0;
    return i(a) - i(b);
  }
  if (field?.type === "number" || field?.type === "integer") return Number(a) - Number(b);
  if (field?.type === "boolean") return Number(b === true) - Number(a === true);
  return String(a).localeCompare(String(b));
}

/** The rows a view shows (with their index in the Excel table), filtered and sorted. */
export function apply(rows: Row[], view: TableView, fields: Field[]) {
  const q = view.search?.trim().toLowerCase();
  const kept = rows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => view.filters.every((f) => matches(row, f, fields)) && (!q || fields.some((f) => listOf(row[f.name]).join(" ").toLowerCase().includes(q) || String(row[f.name] ?? "").toLowerCase().includes(q))));
  return kept.sort((a, b) => {
    for (const s of view.sorts) {
      const c = compare(a.row[s.field], b.row[s.field], fields.find((f) => f.name === s.field));
      if (c) return s.desc ? -c : c;
    }
    return a.index - b.index;
  });
}

/** Rows split by the values of an option field (options' order, then no value); one group without it. */
export function groupRows<T extends { row: Row }>(items: T[], field: Field | undefined) {
  if (!field) return [{ key: "", label: "", color: undefined, items }];
  const keys = [...(field.options ?? []).map((o) => o.value), ""];
  return keys
    .map((key) => ({ key, label: key || "No value", color: field.options?.find((o) => o.value === key)?.color, items: items.filter(({ row }) => (listOf(row[field.name])[0] ?? "") === key) }))
    .filter((g) => g.items.length);
}

export function filterLabel(f: Filter, fields: Field[]) {
  const field = fields.find((x) => x.name === f.field);
  const value = f.values?.length ? f.values.join(", ") : f.value;
  return `${field?.label || field?.name || f.field} ${OPS[f.op]}${value && !NO_VALUE.includes(f.op) ? ` ${value}` : ""}`;
}

/** The view in one line, for the AI. */
export const viewText = (view: TableView, fields: Field[]) =>
  [
    view.search && `search "${view.search}"`,
    view.filters.length && `filters: ${view.filters.map((f) => filterLabel(f, fields)).join("; ")}`,
    view.sorts.length && `sorted by ${view.sorts.map((s) => `${s.field}${s.desc ? " desc" : ""}`).join(", ")}`,
    view.group && `grouped by ${view.group}`,
  ]
    .filter(Boolean)
    .join(", ") || "no filter";
