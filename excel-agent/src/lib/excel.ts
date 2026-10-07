/**
 * The rows of a file, typed by its schema. Where the file lives (local disk, SharePoint…) is the
 * storage layer's business (src/storage): this class validates values, keeps the rows, sends the
 * changes to the backend and persists them (debounced).
 */
import type { ColumnSpec, TableBackend, Values } from "@/storage/types";
import { CELL_COLORS } from "./colors";
import type { Field, Row, Schema } from "./schema";
import { coerce, labelOf } from "./schema";

export type SaveState = { status: "saved" | "saving" | "error"; error?: string; at?: number };

const WIDTH: Partial<Record<Field["type"], number>> = { text: 48, string: 24, date: 12, boolean: 8, image: 28 };

export class ExcelFile {
  rows: Row[] = [];
  /** Incremented on each change (for React subscriptions). */
  version = 0;
  save: SaveState = { status: "saved" };
  private timer?: ReturnType<typeof setTimeout>;
  private writing = Promise.resolve();
  private listeners = new Set<() => void>();

  private constructor(
    readonly backend: TableBackend,
    private schema: Schema,
  ) {}

  /** Opens the file: sheet / table / missing columns of the schema created when needed. */
  static async open(backend: TableBackend, schema: Schema) {
    const excel = new ExcelFile(backend, schema);
    await excel.load();
    return excel;
  }

  private async load() {
    const columns: ColumnSpec[] = [
      { name: "id", width: 38 },
      ...this.schema.fields.flatMap((f): ColumnSpec[] => [
        {
          name: f.name,
          note: labelOf(f),
          width: WIDTH[f.type] ?? 16,
          date: f.type === "date",
          rich: f.type === "text",
          list: f.type === "option" && !f.multiple ? f.options?.map((o) => o.value) : undefined,
          multiple: f.multiple,
          colors: f.type === "option" ? f.options?.map((o) => ({ value: o.value, ...CELL_COLORS[o.color] })) : undefined,
        },
        // The picture itself, next to its link.
        ...(f.type === "image" ? [{ name: `${f.name}_image`, note: `${labelOf(f)} (picture)`, width: 14, imageOf: f.name }] : []),
      ]),
    ];
    const records = await this.backend.open(this.schema.sheet, columns);
    this.rows = records.map((record) => {
      const row: Row = { id: String(record.id) };
      for (const field of this.schema.fields) {
        const raw = record[field.name] ?? null;
        try {
          row[field.name] = coerce(field, raw);
        } catch {
          row[field.name] = raw; // kept as is, shown as invalid
        }
      }
      return row;
    });
  }

  subscribe(listener: () => void) {
    this.listeners.add(listener);
    return () => void this.listeners.delete(listener);
  }
  private emit() {
    this.version++;
    this.listeners.forEach((l) => l());
  }

  private check(values: Record<string, unknown>) {
    const out: Record<string, unknown> = {};
    for (const [name, value] of Object.entries(values)) {
      if (name === "id") continue;
      const field = this.schema.fields.find((f) => f.name === name);
      if (!field) throw new Error(`Unknown field "${name}" (fields: ${this.schema.fields.map((f) => f.name).join(", ")})`);
      out[name] = coerce(field, value);
    }
    return out;
  }

  /** Values as cells: options joined with "; ", empty arrays as empty cells. */
  private cells(values: Record<string, unknown>): Values {
    return Object.fromEntries(
      Object.entries(values).map(([name, value]) => [name, Array.isArray(value) ? value.join("; ") || null : ((value ?? null) as Values[string])]),
    );
  }

  /** Changes values of existing rows; returns the previous values (to undo). */
  update(changes: { id: string; values: Record<string, unknown> }[]) {
    const checked = changes.map((c) => ({ id: c.id, values: this.check(c.values) }));
    const before = checked.map(({ id, values }) => {
      const row = this.rows.find((r) => r.id === id);
      if (!row) throw new Error(`No row with id ${id}`);
      return { id, values: Object.fromEntries(Object.keys(values).map((k) => [k, row[k]])) };
    });
    this.backend.update(checked.map(({ id, values }) => ({ id, values: this.cells(values) })));
    for (const { id, values } of checked) {
      const index = this.rows.findIndex((r) => r.id === id);
      this.rows[index] = { ...this.rows[index], ...values };
    }
    this.changed();
    return before;
  }

  /** Adds rows at the end (keeping a given unused id, e.g. to restore deleted rows); returns them. */
  insert(values: Record<string, unknown>[]) {
    const required = this.schema.fields.filter((f) => f.required);
    const rows = values.map((v) => {
      const checked = this.check(v);
      const missing = required.filter((f) => checked[f.name] === undefined || checked[f.name] === null || checked[f.name] === "");
      if (missing.length) throw new Error(`Missing required field(s): ${missing.map((f) => f.name).join(", ")}`);
      const free = typeof v.id === "string" && v.id && !this.rows.some((r) => r.id === v.id);
      const row: Row = { id: free ? (v.id as string) : crypto.randomUUID() };
      for (const field of this.schema.fields) row[field.name] = field.name in checked ? checked[field.name] : coerce(field, null);
      return row;
    });
    this.backend.append(rows.map((row) => this.cells(row)));
    this.rows.push(...rows);
    this.changed();
    return rows;
  }

  /** Deletes rows; returns them (to undo). */
  remove(ids: string[]) {
    const removed = this.rows.filter((r) => ids.includes(r.id));
    if (removed.length !== ids.length) throw new Error(`Unknown id(s): ${ids.filter((id) => !removed.some((r) => r.id === id)).join(", ")}`);
    this.backend.remove(ids);
    this.rows = this.rows.filter((r) => !ids.includes(r.id));
    this.changed();
    return removed;
  }

  private changed() {
    this.rows = [...this.rows];
    this.save = { status: "saving" };
    this.emit();
    clearTimeout(this.timer);
    this.timer = setTimeout(() => void this.persist(), 400);
  }

  /** Sends the pending changes (serialized). */
  persist() {
    this.writing = this.writing.then(async () => {
      try {
        await this.backend.flush();
        this.save = { status: "saved", at: Date.now() };
      } catch (error) {
        this.save = { status: "error", error: error instanceof Error ? error.message : String(error) };
      }
      this.emit();
    });
    return this.writing;
  }

  /** Reloads when the data was changed outside (Excel, a colleague…) or after a failed save; not while saving. */
  async reloadIfChanged() {
    if (this.save.status === "saving") return false;
    if (this.save.status === "saved" && !(await this.backend.changedOutside())) return false;
    await this.load();
    this.save = { status: "saved", at: Date.now() };
    this.emit();
    return true;
  }
}
