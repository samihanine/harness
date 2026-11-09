/**
 * The Excel screen's agent: rows read and changed by condition (no id), columns of the 4 types, the
 * view shown. Changes are written to the file at once (sync.ts); the last one can be undone.
 */
import { z } from "zod";
import { tool, type Agent } from "@/agent/loop";
import { withContexts } from "@/agent/context";
import { COLORS } from "@/components/badge";
import { EXCEL_INSTRUCTIONS } from "@/prompts/excel";
import type { ExcelTable } from "./sync";
import { COLUMN_TYPES, DATE_RANGES, columnOf, filterText, rowMatches, viewRows, type Column, type DateRange, type Layout, type Row, type TableEntry, type View, type Where } from "./table";

const condition = z.union([z.string(), z.array(z.string()), z.object({ contains: z.string().optional(), empty: z.boolean().optional(), before: z.string().optional(), after: z.string().optional(), not: z.union([z.string(), z.array(z.string())]).optional() })]);
const whereArg = z.record(z.string(), condition).optional().describe('Per column: "value", ["a","b"], {"contains":"x"}, {"empty":true}, {"before":"yyyy-mm-dd"}, {"after":…}, {"not":…}');
const rowsArg = z.array(z.number()).optional().describe("Row numbers as shown (1 = first row)");
const optionArg = z.preprocess((o) => (typeof o === "string" ? { value: o } : o), z.object({ value: z.string(), color: z.enum(COLORS).optional() }));

const short = (v: string, n = 80) => (v.length > n ? `${v.slice(0, n)}…` : v).replace(/\n/g, " ");
const columnText = (c: Column) => `"${c.name}" (${c.type}${c.options ? `: ${c.options.map((o) => `${o.value} [${o.color}]`).join(", ")}` : ""})`;
const rowText = (columns: Column[], row: Row, index: number) => `${index + 1}. ${columns.map((c) => `${c.name}: ${short(row[c.name] ?? "") || "—"}`).join(" | ")}`;

/** Rows matching a condition and/or row numbers (positions in the table). */
function select(table: ExcelTable, columns: Column[], where?: Where, rows?: number[]) {
  const all = table.data?.rows ?? [];
  if (!where && !rows) throw new Error('Say which rows: "where" (a condition on columns) or "rows" (row numbers).');
  const out = all.map((row, index) => ({ row, index })).filter(({ row, index }) => (!rows || rows.includes(index + 1)) && (!where || rowMatches(row, where, columns)));
  if (!out.length) throw new Error("No row matches: check the values with read_rows (they must be written as in the table).");
  return out;
}

export function excelAgent(get: () => TableEntry, table: () => ExcelTable, view: () => View, setView: (v: View) => void): Agent {
  /** The last change, as what puts it back. */
  let undo: (() => Promise<unknown>) | undefined;
  const cols = () => get().columns;
  return {
    instructions: EXCEL_INSTRUCTIONS,
    context: async () => withContexts([`Excel table "${get().title}" (file ${get().excel.name}, table ${get().excel.table}).`], [get().contextIds]),
    state: async () => {
      const t = table();
      if (!t.data) return "The table is being read.";
      const v = view();
      const shown = viewRows(t.data.rows, v);
      return [
        `Columns: ${cols().map(columnText).join("; ")}`,
        `View "${v.name}" (${v.layout}${v.group ? `, grouped by ${v.group}` : ""}${v.sort ? `, sorted by ${v.sort.column}${v.sort.desc ? " desc" : ""}` : ""})${v.filters.length ? `, filters: ${v.filters.map(filterText).join("; ")}` : ""}${v.search ? `, search "${v.search}"` : ""}: ${shown.length} of ${t.data.rows.length} rows.`,
        ...shown.slice(0, 60).map(({ row, index }) => rowText(cols(), row, index)),
        shown.length > 60 ? `… ${shown.length - 60} more (read_rows).` : "",
      ].filter(Boolean).join("\n");
    },
    tools: [
      tool({
        name: "read_rows",
        description: "Reads rows (all their columns, long texts whole) matching a condition, or by number; none: all.",
        args: z.object({ where: whereArg, rows: rowsArg, limit: z.number().optional() }),
        readOnly: true,
        run: async ({ where, rows, limit }) => {
          const all = (table().data?.rows ?? []).map((row, index) => ({ row, index }));
          const found = where || rows ? select(table(), cols(), where as Where, rows) : all;
          return found.slice(0, limit ?? 100).map(({ row, index }) => `${index + 1}. ${JSON.stringify(row)}`).join("\n") + (found.length > (limit ?? 100) ? `\n… ${found.length - (limit ?? 100)} more` : "");
        },
      }),
      tool({
        name: "update_rows",
        description: "Sets values on every row matching the condition (or the given row numbers).",
        args: z.object({ where: whereArg, rows: rowsArg, set: z.record(z.string(), z.string()).describe("Column → new value") }),
        run: async ({ where, rows, set }) => {
          const found = select(table(), cols(), where as Where, rows);
          const before = found.map(({ index, row }) => ({ index, values: Object.fromEntries(Object.keys(set).map((k) => [columnOf(cols(), k)?.name ?? k, row[columnOf(cols(), k)?.name ?? k] ?? ""])) }));
          await table().update(found.map(({ index }) => ({ index, values: set })));
          undo = () => table().update(before);
          return `${found.length} row(s) changed: ${found.map((f) => f.index + 1).join(", ")}.`;
        },
      }),
      tool({
        name: "add_rows",
        description: "Adds rows at the end of the table (each: column → value; missing columns stay empty).",
        args: z.object({ rows: z.array(z.record(z.string(), z.string())).min(1) }),
        run: async ({ rows }) => {
          const at = await table().add(rows);
          undo = () => table().remove(at);
          return `${rows.length} row(s) added (rows ${at.map((i) => i + 1).join(", ")}).`;
        },
      }),
      tool({
        name: "delete_rows",
        description: "Deletes every row matching the condition (or the given row numbers).",
        args: z.object({ where: whereArg, rows: rowsArg }),
        run: async ({ where, rows }) => {
          const found = select(table(), cols(), where as Where, rows);
          await table().remove(found.map((f) => f.index));
          undo = () => table().add(found.map((f) => f.row));
          return `${found.length} row(s) deleted.`;
        },
      }),
      tool({
        name: "add_column",
        description: "Adds a column at the end (or after another): type text, longtext, date or option (with its options and colors).",
        args: z.object({ name: z.string().min(1), type: z.enum(COLUMN_TYPES as [string, ...string[]]), options: z.array(optionArg).optional(), after: z.string().optional() }),
        run: async ({ name, type, options, after }) => {
          if (columnOf(cols(), name)) throw new Error(`A column "${name}" exists.`);
          const column: Column = { name, type: type as Column["type"], ...(type === "option" && { options: (options ?? []).map((o, i) => ({ value: o.value, color: o.color ?? COLORS[(i + 1) % COLORS.length] })) }) };
          const list = [...cols()];
          const at = after ? list.findIndex((c) => c === columnOf(list, after)) + 1 : list.length;
          list.splice(at > 0 ? at : list.length, 0, column);
          await table().setColumns(list);
          return `Column "${name}" added and written to the file.`;
        },
      }),
      tool({
        name: "update_column",
        description: "Changes a column: rename, type, options (the whole list, each {value, color}), position (after another column).",
        args: z.object({ name: z.string(), rename: z.string().optional(), type: z.enum(COLUMN_TYPES as [string, ...string[]]).optional(), options: z.array(optionArg).optional(), after: z.string().optional() }),
        run: async ({ name, rename, type, options, after }) => {
          const c = columnOf(cols(), name);
          if (!c) throw new Error(`No column "${name}".`);
          const next: Column = { ...c, ...(rename && { name: rename, was: c.was ?? c.name }), ...(type && { type: type as Column["type"] }) };
          if (options) next.options = options.map((o, i) => ({ value: o.value, color: o.color ?? c.options?.find((x) => x.value === o.value)?.color ?? COLORS[(i + 1) % COLORS.length] }));
          if (next.type === "option" && !next.options?.length) next.options = [...new Set((table().data?.rows ?? []).map((r) => r[c.name]).filter(Boolean))].slice(0, 12).map((value, i) => ({ value, color: COLORS[(i + 1) % COLORS.length] }));
          let list = cols().map((x) => (x === c ? next : x));
          if (after) {
            list = list.filter((x) => x !== next);
            list.splice(list.findIndex((x) => x === columnOf(list, after)) + 1, 0, next);
          }
          await table().setColumns(list);
          return `Column "${name}" updated and written to the file.`;
        },
      }),
      tool({
        name: "remove_column",
        description: "Deletes a column and its values from the file (only when the user asks).",
        args: z.object({ name: z.string() }),
        run: async ({ name }) => {
          const c = columnOf(cols(), name);
          if (!c) throw new Error(`No column "${name}".`);
          await table().setColumns(cols().filter((x) => x !== c), [c.was ?? c.name]);
          return `Column "${name}" deleted.`;
        },
      }),
      tool({
        name: "set_view",
        description: "Changes what the screen shows (not the data): filters (replace the view's), layout, grouping (board: an option column), sort, search.",
        args: z.object({
          filters: z.array(z.object({ column: z.string(), values: z.array(z.string()).optional(), text: z.string().optional(), date: z.enum(Object.keys(DATE_RANGES) as [DateRange, ...DateRange[]]).optional() })).optional(),
          layout: z.enum(["table", "board", "list"]).optional(),
          group: z.string().optional(),
          sort: z.object({ column: z.string(), desc: z.boolean().optional() }).optional(),
          search: z.string().optional(),
        }),
        readOnly: true,
        run: async ({ filters, layout, group, sort, search }) => {
          const v = view();
          const name = (n: string) => columnOf(cols(), n)?.name ?? n;
          setView({ ...v, ...(filters && { filters: filters.map((f) => ({ ...f, column: name(f.column) })) }), ...(layout && { layout: layout as Layout }), ...(group && { group: name(group) }), ...(sort && { sort: { ...sort, column: name(sort.column) } }), ...(search !== undefined && { search }) });
          return "View changed.";
        },
      }),
      tool({
        name: "undo_last_change",
        description: "Puts back the rows as they were before your last change (only when the user asks).",
        args: z.object({}),
        run: async () => {
          if (!undo) throw new Error("Nothing to undo.");
          await undo();
          undo = undefined;
          return "Last change undone.";
        },
      }),
    ],
  };
}
