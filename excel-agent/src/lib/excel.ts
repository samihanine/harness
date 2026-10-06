/**
 * A local .xlsx file seen as a table of rows (header on line 1 of the schema's sheet, `id` first).
 *
 * The workbook stays in memory: each change only touches the cells concerned, then the file is
 * written back (debounced). Styles, other sheets and formulas elsewhere are kept as ExcelJS
 * reads them.
 */
import ExcelJS from "exceljs";
import type { Field, Row, Schema } from "./schema";
import { coerce, labelOf } from "./schema";

export type SaveState = { status: "saved" | "saving" | "error"; error?: string; at?: number };

const WIDTH: Partial<Record<Field["type"], number>> = { text: 48, string: 24, date: 12, boolean: 8, image: 28 };
const VALIDATED_ROWS = 5000;

/** Sheet-level data validations (not typed by ExcelJS); ranges do not create rows. */
type Validations = { add(range: string, validation: ExcelJS.DataValidation): void };

export class ExcelFile {
  rows: Row[] = [];
  /** Incremented on each change (for React subscriptions). */
  version = 0;
  save: SaveState = { status: "saved" };
  private columns = new Map<string, number>();
  private timer?: ReturnType<typeof setTimeout>;
  private writing = Promise.resolve();
  private lastModified = 0;
  private listeners = new Set<() => void>();

  private constructor(
    private handle: FileSystemFileHandle,
    private schema: Schema,
    private workbook: ExcelJS.Workbook,
  ) {}

  /** Opens (and creates when empty) the file; adds the sheet / missing columns of the schema. */
  static async open(handle: FileSystemFileHandle, schema: Schema) {
    const file = await handle.getFile();
    const workbook = new ExcelJS.Workbook();
    if (file.size > 0) await workbook.xlsx.load(await file.arrayBuffer());
    const excel = new ExcelFile(handle, schema, workbook);
    excel.lastModified = file.lastModified;
    if (excel.prepare() || file.size === 0) await excel.write();
    return excel;
  }

  subscribe(listener: () => void) {
    this.listeners.add(listener);
    return () => void this.listeners.delete(listener);
  }
  private emit() {
    this.version++;
    this.listeners.forEach((l) => l());
  }

  private get sheet() {
    return this.workbook.getWorksheet(this.schema.sheet)!;
  }

  /** Sheet, header and validations; reads the rows. Returns true when the structure changed. */
  private prepare() {
    let changed = false;
    let sheet = this.workbook.getWorksheet(this.schema.sheet);
    if (!sheet) {
      sheet = this.workbook.addWorksheet(this.schema.sheet, { views: [{ state: "frozen", ySplit: 1 }] });
      changed = true;
    }
    const header = sheet.getRow(1);
    this.columns.clear();
    header.eachCell((cell, col) => {
      const name = String(cell.value ?? "").trim();
      if (name) this.columns.set(name, col);
    });
    for (const name of ["id", ...this.schema.fields.map((f) => f.name)]) {
      if (this.columns.has(name)) continue;
      const col = Math.max(0, ...this.columns.values()) + 1;
      const cell = header.getCell(col);
      cell.value = name;
      cell.font = { bold: true };
      const field = this.schema.fields.find((f) => f.name === name);
      sheet.getColumn(col).width = field ? (WIDTH[field.type] ?? 16) : 38;
      if (field) cell.note = labelOf(field);
      this.columns.set(name, col);
      changed = true;
    }
    if (changed) this.validate();

    // Rows: every line below the header with content; rows without id get one.
    this.rows = [];
    for (let r = 2; r <= sheet.rowCount; r++) {
      const line = sheet.getRow(r);
      if (!line.hasValues) continue;
      let id = text(line.getCell(this.columns.get("id")!).value);
      if (!id) {
        id = crypto.randomUUID();
        line.getCell(this.columns.get("id")!).value = id;
        changed = true;
      }
      const row: Row = { id };
      for (const field of this.schema.fields) {
        const raw = cellValue(line.getCell(this.columns.get(field.name)!).value);
        try {
          row[field.name] = coerce(field, raw);
        } catch {
          row[field.name] = raw; // kept as is, shown as invalid
        }
      }
      this.rows.push(row);
    }
    return changed;
  }

  /** Dropdown lists for single option fields, wrapped text for long text. */
  private validate() {
    const sheet = this.sheet;
    for (const field of this.schema.fields) {
      const col = this.columns.get(field.name)!;
      if (field.type === "option" && !field.multiple && field.options?.length) {
        const list = `"${field.options.map((o) => o.value.replace(/[",]/g, "")).join(",")}"`;
        const letter = sheet.getColumn(col).letter;
        (sheet as unknown as { dataValidations: Validations }).dataValidations.add(`${letter}2:${letter}${VALIDATED_ROWS}`, {
          type: "list",
          allowBlank: true,
          formulae: [list],
        });
      }
      if (field.type === "text") sheet.getColumn(col).alignment = { wrapText: true, vertical: "top" };
    }
  }

  /** Last line holding values (rowCount may include empty formatted lines). */
  private lastLine() {
    let line = this.sheet.rowCount;
    while (line > 1 && !this.sheet.getRow(line).hasValues) line--;
    return line;
  }

  private lineOf(id: string) {
    const col = this.columns.get("id")!;
    const sheet = this.sheet;
    for (let r = 2; r <= sheet.rowCount; r++) if (text(sheet.getRow(r).getCell(col).value) === id) return r;
    throw new Error(`No row with id ${id}`);
  }

  private writeCells(line: number, row: Row) {
    const excelRow = this.sheet.getRow(line);
    excelRow.getCell(this.columns.get("id")!).value = row.id;
    for (const field of this.schema.fields) {
      if (!(field.name in row)) continue;
      excelRow.getCell(this.columns.get(field.name)!).value = toCell(field, row[field.name]);
    }
    excelRow.commit();
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

  /** Changes values of existing rows; returns the previous values (to undo). */
  update(changes: { id: string; values: Record<string, unknown> }[]) {
    const checked = changes.map((c) => ({ id: c.id, values: this.check(c.values) }));
    const before = checked.map(({ id, values }) => {
      const row = this.rows.find((r) => r.id === id);
      if (!row) throw new Error(`No row with id ${id}`);
      return { id, values: Object.fromEntries(Object.keys(values).map((k) => [k, row[k]])) };
    });
    for (const { id, values } of checked) {
      const index = this.rows.findIndex((r) => r.id === id);
      this.rows[index] = { ...this.rows[index], ...values };
      this.writeCells(this.lineOf(id), this.rows[index]);
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
    let line = this.lastLine();
    for (const row of rows) this.writeCells(++line, row);
    this.rows.push(...rows);
    this.changed();
    return rows;
  }

  /** Deletes rows; returns them (to undo). */
  remove(ids: string[]) {
    const removed = this.rows.filter((r) => ids.includes(r.id));
    if (removed.length !== ids.length) throw new Error(`Unknown id(s): ${ids.filter((id) => !removed.some((r) => r.id === id)).join(", ")}`);
    for (const id of ids) this.sheet.spliceRows(this.lineOf(id), 1);
    this.rows = this.rows.filter((r) => !ids.includes(r.id));
    this.changed();
    return removed;
  }

  private changed() {
    this.rows = [...this.rows];
    this.save = { status: "saving" };
    this.emit();
    clearTimeout(this.timer);
    this.timer = setTimeout(() => void this.write(), 400);
  }

  /** Writes the workbook to the file (serialized). */
  write() {
    this.writing = this.writing.then(async () => {
      try {
        const buffer = await this.workbook.xlsx.writeBuffer();
        const stream = await this.handle.createWritable();
        await stream.write(buffer);
        await stream.close();
        this.lastModified = (await this.handle.getFile()).lastModified;
        this.save = { status: "saved", at: Date.now() };
      } catch (error) {
        this.save = {
          status: "error",
          error: `Could not write the file${error instanceof Error ? `: ${error.message}` : ""}. Is it open in Excel?`,
        };
      }
      this.emit();
    });
    return this.writing;
  }

  /** Reloads the file when it was changed outside (Excel…) and nothing is pending here. */
  async reloadIfChanged() {
    if (this.save.status !== "saved") return false;
    const file = await this.handle.getFile();
    if (file.lastModified <= this.lastModified) return false;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    this.workbook = workbook;
    this.lastModified = file.lastModified;
    if (this.prepare()) await this.write();
    this.emit();
    return true;
  }
}

const text = (value: ExcelJS.CellValue) => String(cellValue(value) ?? "").trim();

/** Plain value of a cell (rich text, formulas, hyperlinks…). */
function cellValue(value: ExcelJS.CellValue): unknown {
  if (value === null || value === undefined) return null;
  if (value instanceof Date || typeof value !== "object") return value;
  if ("richText" in value) return value.richText.map((r) => r.text).join("");
  if ("result" in value) return cellValue(value.result as ExcelJS.CellValue);
  if ("text" in value) return value.text;
  if ("error" in value) return null;
  return String(value);
}

function toCell(field: Field, value: unknown): ExcelJS.CellValue {
  if (value === null || value === undefined) return null;
  if (field.type === "option" && Array.isArray(value)) return value.join("; ") || null;
  if (field.type === "date") return new Date(`${value}T00:00:00Z`);
  return value as ExcelJS.CellValue;
}
