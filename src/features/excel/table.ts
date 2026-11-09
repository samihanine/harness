/**
 * An Excel table as a simple database: 4 column types (text, long text, date, option with colors),
 * saved views (table, board, list) with simple filters, rows read and written as text. The columns'
 * types and the views are kept with the entry and in a hidden sheet of the file (services/excel.ts),
 * so the file carries them to every user.
 */
import { COLORS, type Color } from "@/components/badge";
import type { Cell, ExcelRef } from "@/services/excel";

export type ColumnType = "text" | "longtext" | "date" | "option";
export const COLUMN_TYPES: ColumnType[] = ["text", "longtext", "date", "option"];
export const TYPE_LABELS: Record<ColumnType, string> = { text: "Text", longtext: "Long text", date: "Date", option: "Option" };

export type Option = { value: string; color: Color };
/** `was`: its name in the file before a rename (until the next sync). */
export type Column = { name: string; type: ColumnType; options?: Option[]; was?: string };

/** A filter: an option column keeps the ticked values; a text column, the rows containing a text; a date column, a period. */
export type DateRange = "past" | "today" | "week" | "month" | "future" | "empty";
export const DATE_RANGES: Record<DateRange, string> = { past: "overdue", today: "today", week: "next 7 days", month: "next 30 days", future: "later", empty: "no date" };
export type Filter = { column: string; values?: string[]; text?: string; date?: DateRange };
export type Layout = "table" | "board" | "list";
export type View = { id: string; name: string; layout: Layout; filters: Filter[]; sort?: { column: string; desc?: boolean }; group?: string; hidden?: string[]; search?: string };

export type TableEntry = {
  id: string;
  title: string;
  excel: ExcelRef;
  columns: Column[];
  views: View[];
  /** Columns deleted in the app, deleted from the file at the next sync. */
  removed?: string[];
  /** When the columns or views last changed here (the file's copy wins when newer). */
  changedAt?: number;
  contextIds: string[];
};
declare module "@/lib/store" {
  interface Collections {
    tables: TableEntry;
  }
}

/** What the hidden sheet of the file keeps. */
export type AppData = { columns: Column[]; views: View[]; at: number };

export type Row = Record<string, string>;
/** The table as read: the file's header (its order), the rows as text, and their cells as the file has them. */
export type TableData = { header: string[]; rows: Row[]; cells: Cell[][] };

export const DEFAULT_VIEW: View = { id: "all", name: "All", layout: "table", filters: [] };

/* ------------------------------------ Cells ----------------------------------- */

const pad = (n: number) => String(n).padStart(2, "0");
const iso = (d: Date) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
/** An Excel date (days since 1899-12-30) as yyyy-mm-dd. */
const fromSerial = (n: number) => iso(new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86_400_000));
export const today = () => {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
};
export const addDays = (date: string, days: number) => iso(new Date(Date.parse(`${date}T00:00:00Z`) + days * 86_400_000));

/** A cell as the app's text (a date column: yyyy-mm-dd). */
export function toText(column: Column | undefined, v: Cell): string {
  if (v === null || v === undefined) return "";
  if (column?.type === "date") {
    if (typeof v === "number") return fromSerial(v);
    const s = String(v).trim();
    const t = Date.parse(s);
    return /^\d{4}-\d{2}-\d{2}$/.test(s) || Number.isNaN(t) ? s : iso(new Date(t - new Date(t).getTimezoneOffset() * 60_000));
  }
  return String(v);
}

/** A value as written to the file (Excel reads yyyy-mm-dd as a date). */
export const toCell = (v: string | undefined): Cell => (v === undefined || v === "" ? "" : v);

/** The value of a row for a column, checked: an option not in the list, a date not readable → an error. */
export function checkValue(column: Column, value: unknown): string {
  const v = value === null || value === undefined ? "" : String(value).trim();
  if (!v) return "";
  if (column.type === "date") {
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    const t = Date.parse(v);
    if (Number.isNaN(t)) throw new Error(`"${v}" is not a date (column ${column.name}: write yyyy-mm-dd).`);
    return iso(new Date(t - new Date(t).getTimezoneOffset() * 60_000));
  }
  if (column.type === "option") {
    const found = column.options?.find((o) => o.value.toLowerCase() === v.toLowerCase());
    if (!found) throw new Error(`"${v}" is not an option of ${column.name}: ${column.options?.map((o) => `"${o.value}"`).join(", ") || "none"} (add the option to the column first).`);
    return found.value;
  }
  return String(value);
}

/** The table read: rows as text by column name (the file's columns the app does not know are kept as text too). */
export function readData(columns: Column[], header: string[], cells: Cell[][]): TableData {
  const of = (name: string) => columns.find((c) => fileName(c, header) === name);
  return { header, cells, rows: cells.map((r) => Object.fromEntries(header.map((h, i) => [of(h)?.name ?? h, toText(of(h), r[i])]))) };
}

/** A row's cells in the file's column order: the given values where set, the file's own cells elsewhere. */
export function rowCells(columns: Column[], header: string[], old: Cell[] | undefined, values: Row) {
  return header.map((h, i) => {
    const c = columns.find((x) => fileName(x, header) === h);
    const name = c?.name ?? h;
    return name in values ? toCell(values[name]) : (old?.[i] ?? "");
  });
}

/* ------------------------------------ Columns ---------------------------------- */

/** The columns of a file never seen: a type guessed from each column's values (options get colors). */
export function guessColumns(header: string[], cells: Cell[][]): Column[] {
  return header.map((name, i) => {
    const values = cells.map((r) => r[i]).filter((v) => v !== null && v !== "" && v !== undefined);
    const texts = values.map(String);
    const distinct = [...new Set(texts)];
    const dateLike = values.length > 0 && values.every((v) => (typeof v === "number" && v > 20_000 && v < 80_000 && /date|due|day|deadline|échéance|jour/i.test(name)) || (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v)));
    if (dateLike) return { name, type: "date" };
    if (texts.some((t) => t.length > 80 || t.includes("\n"))) return { name, type: "longtext" };
    // Few short values repeated, or a column named like a choice (status, priority, owner…).
    const choiceName = /status|state|priority|type|category|owner|team|stage|phase|level|statut|état|priorité|catégorie|équipe|responsable|niveau/i.test(name);
    if (distinct.length > 0 && distinct.length <= 8 && texts.every((t) => t.length <= 30) && (choiceName || distinct.length <= Math.max(2, values.length * 0.6))) return { name, type: "option", options: distinct.map((value, k) => ({ value, color: COLORS[(k + 1) % COLORS.length] })) };
    return { name, type: "text" };
  });
}

/** The name a column has in the file now (its former name until a rename is synced), none when not there yet. */
export const fileName = (c: Column, header: string[]) => (c.was && header.includes(c.was) ? c.was : header.includes(c.name) ? c.name : undefined);

/**
 * The app's columns against the file's header: known ones kept (in the app's order), the file's new
 * ones added at the end (types guessed). `pending`: what the next sync changes in the file (columns to
 * add, to delete, an order or names to rewrite).
 */
export function reconcile(columns: Column[], header: string[], cells: Cell[][], removed: string[] = []) {
  const known = (h: string) => columns.some((c) => fileName(c, header) === h);
  const added = guessColumns(header, cells).filter((c) => !known(c.name) && !removed.includes(c.name) && !columns.some((x) => x.name === c.name));
  const next = [...columns, ...added];
  const present = next.filter((c) => fileName(c, header));
  const pending = {
    add: next.filter((c) => !fileName(c, header)).map((c) => c.name),
    remove: removed.filter((r) => header.includes(r)),
    reorder: JSON.stringify(header.filter((h) => !removed.includes(h))) !== JSON.stringify(present.map((c) => fileName(c, header))) || present.some((c) => fileName(c, header) !== c.name),
  };
  return { columns: next, pending, dirty: pending.add.length > 0 || pending.remove.length > 0 || pending.reorder };
}

/* ------------------------------------ Filters ---------------------------------- */

export function inRange(v: string, range: DateRange) {
  if (range === "empty") return !v;
  if (!v) return false;
  const t = today();
  if (range === "past") return v < t;
  if (range === "today") return v === t;
  if (range === "week") return v >= t && v <= addDays(t, 7);
  if (range === "month") return v >= t && v <= addDays(t, 30);
  return v > t;
}

export function matches(row: Row, f: Filter) {
  const v = row[f.column] ?? "";
  if (f.values) return !f.values.length || f.values.includes(v);
  if (f.date) return inRange(v, f.date);
  if (f.text !== undefined) return !f.text.trim() || v.toLowerCase().includes(f.text.trim().toLowerCase());
  return true;
}

/** The rows of a view (their positions kept): filters, search, sort. */
export function viewRows(rows: Row[], view: View): { row: Row; index: number }[] {
  const q = view.search?.trim().toLowerCase();
  const out = rows.map((row, index) => ({ row, index })).filter(({ row }) => view.filters.every((f) => matches(row, f)) && (!q || Object.values(row).some((v) => v.toLowerCase().includes(q))));
  if (view.sort) {
    const { column, desc } = view.sort;
    out.sort((a, b) => (a.row[column] ?? "").localeCompare(b.row[column] ?? "", undefined, { numeric: true }) * (desc ? -1 : 1));
  }
  return out;
}

/** A filter in words (chips, the AI's state). */
export const filterText = (f: Filter) => (f.values ? `${f.column} is ${f.values.join(" or ") || "anything"}` : f.date ? `${f.column}: ${DATE_RANGES[f.date]}` : `${f.column} contains "${f.text ?? ""}"`);

/* ------------------------------------- Where ----------------------------------- */

/**
 * The AI's condition on rows (no id needed): per column, a value, a list of values, or
 * {contains, empty, before, after}; `rows`: row numbers as shown (1 = first row).
 */
export type Condition = string | string[] | { contains?: string; empty?: boolean; before?: string; after?: string; not?: string | string[] };
export type Where = Record<string, Condition>;

export function rowMatches(row: Row, where: Where, columns: Column[]) {
  return Object.entries(where).every(([name, cond]) => {
    const column = columnOf(columns, name);
    if (!column) throw new Error(`No column "${name}": ${columns.map((c) => `"${c.name}"`).join(", ")}.`);
    const v = (row[column.name] ?? "").trim();
    const same = (x: string) => v.toLowerCase() === String(x).trim().toLowerCase();
    if (typeof cond === "string") return same(cond);
    if (Array.isArray(cond)) return cond.some(same);
    if (cond.empty !== undefined && !v !== cond.empty) return false;
    if (cond.contains && !v.toLowerCase().includes(cond.contains.toLowerCase())) return false;
    if (cond.not !== undefined && (Array.isArray(cond.not) ? cond.not.some(same) : same(cond.not))) return false;
    if (cond.before && !(v && v < cond.before)) return false;
    if (cond.after && !(v && v > cond.after)) return false;
    return true;
  });
}

/** A column by name, the case and spaces aside. */
export const columnOf = (columns: Column[], name: string) => {
  const key = (s: string) => s.toLowerCase().replace(/[\s_]+/g, " ").trim();
  return columns.find((c) => c.name === name) ?? columns.find((c) => key(c.name) === key(name));
};
