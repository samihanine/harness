/**
 * The Excel screen's data tools: reading, finding and changing the rows of a schema's Excel table, and
 * the user's view of them (search, filters, sorts). Its structure is changed by the schema tools.
 */
import { z } from "zod";
import { tool, type Agent } from "./loop";
import { withContexts } from "./context";
import { addRows, deleteRows, idError, updateRows, withRows, withUpdates, withoutRows, type Table } from "@/lib/excel";
import { rowsText } from "@/lib/pbi";
import { TYPE_LABELS, idPrefix, isIdField, type Field } from "@/lib/schema";
import type { SchemaEntry } from "@/lib/store";
import { OPS, matches, viewText, type Filter, type TableView } from "@/lib/table-view";
import type { Write } from "@/lib/use-table";

const cell = z.union([z.string(), z.number(), z.boolean(), z.array(z.string()), z.null()]);
const filterArg = z.object({ field: z.string(), op: z.enum(Object.keys(OPS) as [Filter["op"], ...Filter["op"][]]), value: z.string().optional(), values: z.array(z.string()).optional() });

/** The fields for the AI: name, type, rules, choices. */
export const fieldsText = (fields: Field[]) =>
  fields
    .map((f) => `- ${f.name} (${TYPE_LABELS[f.type]}${f.type === "rating" ? ` 1-${f.max ?? 5}` : ""}${f.required ? ", required" : ""}${f.multiple ? ", several values" : ""})${f.label ? ` label "${f.label}"` : ""}${f.options?.length ? ` choices: ${f.options.map((o) => o.value).join(", ")}` : ""}${f.description ? ` — ${f.description}` : ""}`)
    .join("\n");

export function excelData({ schema, table, shown, checked, view, setView, openRow, write }: {
  schema: SchemaEntry;
  table: Table;
  /** Excel indexes of the rows the view shows, in its order. */
  shown: () => number[];
  checked: number[];
  view: TableView;
  setView: (v: Partial<TableView>) => void;
  openRow: (index: number) => void;
  write: Write;
}): Pick<Agent, "instructions" | "tools" | "state" | "context"> {
  const excel = schema.excel!;
  const idField = table.fields.find(isIdField);
  /** Row given by its id (TASK_001) or its index. */
  const indexOf = (r: { index?: number; id?: string }) => {
    if (r.id === undefined) {
      if (r.index === undefined) throw new Error("Give the row id (or index).");
      return r.index;
    }
    const index = idField ? table.rows.findIndex((x) => String(x[idField.name] ?? "").toLowerCase() === r.id!.toLowerCase()) : -1;
    if (index < 0) throw new Error(`No row with id ${r.id}`);
    return index;
  };
  const pick = (index: number) => ({ index, ...Object.fromEntries(table.fields.map((f) => [f.name, table.rows[index]?.[f.name]])) });
  const flat = (r: Record<string, unknown>) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, Array.isArray(v) ? v.join("; ") : v]));
  const checkFields = (names: string[]) => {
    const unknown = names.filter((n) => !table.fields.some((f) => f.name === n));
    if (unknown.length) throw new Error(`Unknown field(s): ${unknown.join(", ")}. Fields: ${table.fields.map((f) => f.name).join(", ")}.`);
  };
  return {
    instructions: `DATA (Data tab): you read, find, add, change and delete the rows of the Excel table "${schema.title}" (SharePoint / OneDrive), and change what the user sees (search, filters, sorts). Rows have ${idField ? `an id (column "${idField.name}", e.g. ${idPrefix(excel.table)}_001, given by the app; use it to name rows to the user and in tool calls; change one with change_row_id ONLY when the user explicitly asks to change that id) and ` : ""}an "index" (position in the Excel table). Long text holds markdown; choice fields only accept their listed values (several values: a list); ratings are whole numbers; dates are YYYY-MM-DD; links, images and files are URLs.
Do the whole request in one go; for many rows, change them all in one update_rows call.`,
    context: async () =>
      withContexts(
        [`Excel table "${excel.table}" of "${excel.name}" (schema "${schema.title}")`, `Fields:\n${fieldsText(table.fields)}`, schema.form && `Microsoft Form "${schema.form.title}" fills this table.`],
        [schema.contextIds],
      ),
    // What the user sees now (the "list" of this screen), read again before each model call.
    state: async () =>
      [
        checked.length
          ? `THE USER SELECTED ${checked.length} ROW(S): indexes ${checked.join(", ")}. Act on these rows only (read them with read_rows), unless the user clearly asks otherwise.`
          : "No row selected: the request concerns all the rows.",
        `${table.rows.length} rows; the user's view shows ${shown().length} (${viewText(view, table.fields)}).`,
        `First rows of the view:\n${rowsText(shown().slice(0, 8).map(pick).map(flat), 8)}`,
      ].join("\n\n"),
    tools: [
      tool({
        name: "read_rows",
        description: "Reads rows with their index: the selected rows when the user selected some, else those of the user's view; all=true for every row.",
        args: z.object({ all: z.boolean().optional(), offset: z.number().optional(), limit: z.number().optional() }),
        readOnly: true,
        run: async ({ all, offset = 0, limit = 30 }) => {
          const indexes = all ? table.rows.map((_, i) => i) : checked.length ? checked : shown();
          return rowsText(indexes.slice(offset, offset + limit).map(pick).map(flat), limit);
        },
      }),
      tool({
        name: "find_rows",
        description: 'Finds rows (all of them, whatever the view) matching every filter, e.g. [{"field":"status","op":"is","values":["Open"]}]. Ops: contains, is, isNot, empty, notEmpty, gt, lt, before, after.',
        args: z.object({ filters: z.array(filterArg), limit: z.number().optional() }),
        readOnly: true,
        run: async ({ filters, limit = 30 }) => {
          checkFields(filters.map((f) => f.field));
          const found = table.rows.map((row, index) => ({ row, index })).filter(({ row }) => filters.every((f) => matches(row, f as Filter, table.fields)));
          return `${found.length} row(s).\n${rowsText(found.slice(0, limit).map(({ index }) => flat(pick(index))), limit)}`;
        },
      }),
      tool({
        name: "update_rows",
        description: "Changes values of rows (only the given fields; never the id: see change_row_id). Each row by its id (or index).",
        args: z.object({ rows: z.array(z.object({ id: z.string().optional(), index: z.number().optional(), values: z.record(z.string(), cell) })) }),
        run: async ({ rows: given }) => {
          const changes = given.map((r) => ({ index: indexOf(r), values: Object.fromEntries(Object.entries(r.values).filter(([k]) => k !== idField?.name)) }));
          await write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) });
          return `${changes.length} row(s) updated.`;
        },
      }),
      tool({
        name: "add_rows",
        description: "Adds rows at the end of the table (their ids are given automatically). Required fields must be given.",
        args: z.object({ rows: z.array(z.record(z.string(), cell)) }),
        run: async ({ rows: added }) => {
          await write((t) => addRows(excel, t, added), { refetch: true, optimistic: (t) => withRows(t, added) });
          return `${added.length} row(s) added.`;
        },
      }),
      tool({
        name: "delete_rows",
        description: "Deletes rows by id (or index). Only when the user asks to delete them.",
        args: z.object({ ids: z.array(z.string()).optional(), indexes: z.array(z.number()).optional() }),
        run: async ({ ids = [], indexes: given = [] }) => {
          const indexes = [...new Set([...ids.map((id) => indexOf({ id })), ...given])];
          await write(() => deleteRows(excel, indexes), { optimistic: (t) => withoutRows(t, indexes) });
          return `${indexes.length} row(s) deleted.`;
        },
      }),
      tool({
        name: "set_view",
        description: 'Changes what the user sees in the Data tab (only the given parts): search, filters, sorts, and opens a row with all its fields. Filter ops as in find_rows; choice fields: op "is" / "isNot" with "values".',
        args: z.object({
          search: z.string().optional(),
          filters: z.array(filterArg).optional(),
          sorts: z.array(z.object({ field: z.string(), desc: z.boolean().optional() })).optional(),
          open_row: z.string().optional().describe("Id of the row to show with all its fields"),
        }),
        run: async ({ search, filters, sorts, open_row }) => {
          checkFields([...(filters ?? []), ...(sorts ?? [])].map((x) => x.field));
          setView({ ...(search !== undefined && { search }), ...(filters && { filters: filters as Filter[] }), ...(sorts && { sorts }) });
          if (open_row !== undefined) openRow(indexOf({ id: open_row }));
          return "View updated.";
        },
      }),
      ...(idField
        ? [
            tool({
              name: "change_row_id",
              description: "Changes the id of a row. ONLY when the user explicitly asks to change / rename that id. The new id must be unique (case ignored) and not empty.",
              args: z.object({ id: z.string().describe("Current id"), new_id: z.string() }),
              run: async ({ id, new_id }) => {
                const index = indexOf({ id });
                const error = idError(table, index, new_id);
                if (error) throw new Error(error);
                const changes = [{ index, values: { [idField.name]: new_id.trim() } }];
                await write((t) => updateRows(excel, t, changes), { optimistic: (t) => withUpdates(t, changes) });
                return `Id ${id} changed to ${new_id.trim()}.`;
              },
            }),
          ]
        : []),
    ],
  };
}
