/**
 * Local .xlsx (File System Access API + ExcelJS). The workbook stays in memory: changes touch
 * only their cells, then flush() writes the file (an .xlsx is a zip: it is rewritten whole).
 * Styles, other sheets and formulas elsewhere are kept as ExcelJS reads them.
 */
import ExcelJS from "exceljs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { htmlToRuns, isStyled, plainToHtml, runsToHtml, runsToText } from "@/lib/rich";
import type { Run } from "@/lib/rich";
import type { Cell, ColumnSpec, TableBackend, Values } from "./types";
import { imageFormula, rowHeight } from "./types";

type TableModel = {
  name: string;
  tableRef: string;
  autoFilterRef: string;
  columns: { name: string; filterButton?: boolean; totalsRowFunction?: string; totalsRowLabel?: string }[];
};
/** Sheet-level data validations (not typed by ExcelJS), by cell or range address; ranges do not create rows. */
type Validations = { model: Record<string, ExcelJS.DataValidation | undefined>; add(range: string, validation: ExcelJS.DataValidation): void };
type WithTables = { getTables(): { table: TableModel }[] };
type Formatting = { ref: string; rules: ExcelJS.ConditionalFormattingRule[] };

/** Rows covered by dropdowns and option colours. */
const VALIDATED_ROWS = 5000;
const argb = (hex: string) => `FF${hex.slice(1).toUpperCase()}`;
const BORDER: Partial<ExcelJS.Borders> = Object.fromEntries(
  (["top", "left", "bottom", "right"] as const).map((side) => [side, { style: "thin", color: { argb: "FFD4D4D8" } }]),
);
const HEADER = { font: { bold: true, color: { argb: "FFFFFFFF" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } } } as const;

export class LocalTable implements TableBackend {
  readonly pollEvery = 3_000;
  private workbook = new ExcelJS.Workbook();
  private sheetName = "";
  private columns = new Map<string, number>();
  private dates = new Set<string>();
  private rich = new Set<string>();
  private lists: ColumnSpec[] = [];
  private specs: ColumnSpec[] = [];
  private lastModified = 0;
  private dirty = false;

  constructor(private handle: FileSystemFileHandle) {}

  private get sheet() {
    return this.workbook.getWorksheet(this.sheetName)!;
  }

  async open(sheetName: string, specs: ColumnSpec[]) {
    const file = await this.handle.getFile();
    this.workbook = new ExcelJS.Workbook();
    if (file.size > 0) await this.workbook.xlsx.load(await file.arrayBuffer());
    this.lastModified = file.lastModified;
    this.sheetName = sheetName;
    this.dates = new Set(specs.filter((s) => s.date).map((s) => s.name));
    this.rich = new Set(specs.filter((s) => s.rich).map((s) => s.name));
    this.lists = specs.filter((s) => s.list?.length);
    this.specs = specs;
    this.dirty = file.size === 0;

    let sheet = this.workbook.getWorksheet(sheetName);
    if (!sheet) {
      sheet = this.workbook.addWorksheet(sheetName, { views: [{ state: "frozen", ySplit: 1 }] });
      this.dirty = true;
    }
    const header = sheet.getRow(1);
    this.columns.clear();
    header.eachCell((cell, col) => {
      const name = text(cell.value);
      if (name) this.columns.set(name, col);
    });
    const added: ColumnSpec[] = [];
    for (const spec of specs) {
      if (this.columns.has(spec.name)) continue;
      const col = Math.max(0, ...this.columns.values()) + 1;
      const cell = header.getCell(col);
      cell.value = spec.name;
      if (spec.note) cell.note = spec.note;
      sheet.getColumn(col).width = spec.width ?? 16;
      this.columns.set(spec.name, col);
      added.push(spec);
    }
    if (added.length) {
      this.dirty = true;
      for (const spec of specs) this.format(spec);
    }
    // Rows styled by an older version (or by hand): restyle them all once.
    const restyle = added.length > 0 || (sheet.rowCount > 1 && sheet.getRow(2).height !== rowHeight(specs));
    if (restyle) this.dirty = true;

    // Rows: every line below the header with content; rows without id get one.
    const rows: Values[] = [];
    const idCol = this.columns.get("id")!;
    for (let r = 2; r <= sheet.rowCount; r++) {
      const line = sheet.getRow(r);
      if (!line.hasValues) continue;
      if (restyle) this.decorate(r);
      if (!text(line.getCell(idCol).value)) {
        line.getCell(idCol).value = crypto.randomUUID();
        this.dirty = true;
      }
      rows.push(Object.fromEntries([...this.columns].map(([name, col]) => [name, this.rich.has(name) ? richCell(line.getCell(col).value) : plain(line.getCell(col).value)])));
    }
    if (this.dirty) await this.flush();
    return rows;
  }

  /** Wrapped text (dropdown lists are set on each write, see validations()). */
  private format(spec: ColumnSpec) {
    const col = this.columns.get(spec.name)!;
    if (spec.wrap || spec.rich) this.sheet.getColumn(col).alignment = { wrapText: true, vertical: "top" };
  }

  /**
   * One dropdown range per option column. ExcelJS loads a range as one entry per cell and writes
   * them back as overlapping ranges, which Excel reports as unreadable content: the entries of
   * these columns are replaced by a single range each time.
   */
  private validations() {
    const { model } = (this.sheet as unknown as { dataValidations: Validations }).dataValidations;
    for (const spec of this.lists) {
      const letter = this.sheet.getColumn(this.columns.get(spec.name)!).letter;
      for (const address of Object.keys(model)) if (address.match(/^([A-Z]+)\d/)?.[1] === letter) delete model[address];
      const values = spec.list!.map((v) => v.replace(/[",]/g, ""));
      model[`${letter}2:${letter}${VALIDATED_ROWS}`] = {
        type: "list",
        allowBlank: true,
        formulae: [`"${values.join(",")}"`],
        showErrorMessage: true,
        errorStyle: "stop",
        errorTitle: "Invalid value",
        error: `Choose one of: ${values.join(", ")}`,
      };
    }
  }

  /** Option cells take the colours of their value (replaced on each write, like the dropdowns). */
  private optionColors() {
    const sheet = this.sheet as unknown as { conditionalFormattings: Formatting[] };
    const specs = this.specs.filter((s) => s.colors?.length);
    const refs = new Set(specs.map((s) => `${this.sheet.getColumn(this.columns.get(s.name)!).letter}2:${this.sheet.getColumn(this.columns.get(s.name)!).letter}${VALIDATED_ROWS}`));
    sheet.conditionalFormattings = (sheet.conditionalFormattings ?? []).filter((f) => !refs.has(f.ref));
    for (const spec of specs) {
      const letter = this.sheet.getColumn(this.columns.get(spec.name)!).letter;
      sheet.conditionalFormattings.push({
        ref: `${letter}2:${letter}${VALIDATED_ROWS}`,
        rules: spec.colors!.map(({ value, fill, font }, i) => {
          const style: Partial<ExcelJS.Style> = { fill: { type: "pattern", pattern: "solid", bgColor: { argb: argb(fill) } }, font: { color: { argb: argb(font) } } };
          return spec.multiple
            ? { type: "containsText", operator: "containsText", text: value, style, priority: i + 1 }
            : { type: "cellIs", operator: "equal", formulae: [`"${value.replace(/"/g, '""')}"`], style, priority: i + 1 };
        }) as ExcelJS.ConditionalFormattingRule[],
      });
    }
  }

  private lineOf(id: string) {
    const col = this.columns.get("id")!;
    for (let r = 2; r <= this.sheet.rowCount; r++) if (text(this.sheet.getRow(r).getCell(col).value) === id) return r;
    throw new Error(`No row with id ${id}`);
  }

  /** Last line holding values (rowCount may include empty formatted lines). */
  private lastLine() {
    let line = this.sheet.rowCount;
    while (line > 1 && !this.sheet.getRow(line).hasValues) line--;
    return line;
  }

  /** Height, borders and alignment of a data line, and the IMAGE() formulas of its computed columns. */
  private decorate(line: number) {
    const row = this.sheet.getRow(line);
    row.height = rowHeight(this.specs);
    for (const spec of this.specs) {
      const cell = row.getCell(this.columns.get(spec.name)!);
      cell.border = BORDER;
      cell.alignment = spec.rich || spec.wrap ? { wrapText: true, vertical: "top" } : spec.imageOf ? { horizontal: "center", vertical: "middle" } : { vertical: "middle" };
      if (spec.imageOf) {
        const link = `${this.sheet.getColumn(this.columns.get(spec.imageOf)!).letter}${line}`;
        cell.value = { formula: imageFormula(link, true) } as ExcelJS.CellFormulaValue;
      }
    }
  }

  private write(line: number, values: Values) {
    const row = this.sheet.getRow(line);
    for (const [name, value] of Object.entries(values)) {
      const col = this.columns.get(name);
      if (!col) continue;
      row.getCell(col).value =
        this.dates.has(name) && typeof value === "string" && value
          ? new Date(`${value}T00:00:00Z`)
          : this.rich.has(name) && typeof value === "string"
            ? richValue(value)
            : (value as ExcelJS.CellValue);
    }
    this.decorate(line);
    row.commit();
    this.dirty = true;
  }

  update(changes: { id: string; values: Values }[]) {
    for (const { id, values } of changes) this.write(this.lineOf(id), values);
  }

  append(rows: Values[]) {
    let line = this.lastLine();
    for (const values of rows) this.write(++line, values);
  }

  remove(ids: string[]) {
    for (const id of ids) this.sheet.spliceRows(this.lineOf(id), 1);
    this.dirty = true;
  }

  async flush() {
    if (!this.dirty) return;
    this.syncTable();
    this.validations();
    this.optionColors();
    this.workbook.calcProperties.fullCalcOnLoad = true; // IMAGE() cells have no cached result
    const buffer = fixSheetOrder(new Uint8Array(await this.workbook.xlsx.writeBuffer()));
    try {
      const stream = await this.handle.createWritable();
      await stream.write(buffer as Uint8Array<ArrayBuffer>);
      await stream.close();
    } catch (error) {
      throw new Error(`Could not write the file${error instanceof Error ? `: ${error.message}` : ""}. Is it open in Excel?`);
    }
    this.lastModified = (await this.handle.getFile()).lastModified;
    this.dirty = false;
  }

  async changedOutside() {
    return (await this.handle.getFile()).lastModified > this.lastModified;
  }

  /**
   * Keeps the rows as an Excel table (header + data, from A1), created when missing, so the file
   * can be filtered in Excel and used as a Power BI / Power Query table.
   */
  private syncTable() {
    const sheet = this.sheet;
    const width = Math.max(...this.columns.values());
    const names = Array.from({ length: width }, (_, i) => text(sheet.getCell(1, i + 1).value) || `Column${i + 1}`);
    const ref = `A1:${sheet.getColumn(width).letter}${Math.max(2, this.lastLine())}`;
    const tables = () => (sheet as unknown as WithTables).getTables().map((t) => t.table);
    let table = tables()[0];
    if (table && !/^A1:/.test(table.tableRef)) return; // a table of the user elsewhere: leave the sheet as is
    if (!table) {
      const used = new Set(this.workbook.worksheets.flatMap((ws) => (ws as unknown as WithTables).getTables().map((t) => t.table.name)));
      let name = this.sheetName.replace(/[^\p{L}\p{N}_]/gu, "_").replace(/^(\d)/, "_$1") || "Data";
      for (let i = 2; used.has(name); i++) name = `${name}_${i}`;
      sheet.addTable({ name, ref: "A1", headerRow: true, style: { theme: "TableStyleLight1", showRowStripes: false }, columns: names.map((n) => ({ name: n, filterButton: true })), rows: [] });
      table = tables()[0];
    }
    // Plain table (no stripes): the look comes from the cell styles below and decorate().
    (table as TableModel & { style?: object }).style = { theme: "TableStyleLight1", showRowStripes: false, showColumnStripes: false, showFirstColumn: false, showLastColumn: false };
    const header = sheet.getRow(1);
    header.height = 28;
    for (let col = 1; col <= width; col++) Object.assign(header.getCell(col), { ...HEADER, border: BORDER, alignment: { vertical: "middle" } });
    table.tableRef = ref;
    table.autoFilterRef = ref;
    table.columns = names.map((name, i) => ({
      ...table.columns.find((c) => c.name === name),
      name,
      filterButton: true,
      ...(i === 0 ? { totalsRowLabel: "Total" } : { totalsRowFunction: "none" }),
    }));
  }
}

/**
 * ExcelJS writes <tableParts> before <legacyDrawing> (header notes) in a sheet that has both;
 * the schema wants the reverse and Excel drops the whole sheet when "repairing" the file.
 */
function fixSheetOrder(bytes: Uint8Array) {
  const files = unzipSync(bytes);
  let changed = false;
  for (const path of Object.keys(files).filter((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p))) {
    const xml = strFromU8(files[path]);
    const fixed = xml.replace(/(<tableParts\b[\s\S]*?<\/tableParts>|<tableParts\b[^>]*\/>)(\s*)(<legacyDrawing\b[^>]*\/>)/, "$3$2$1");
    if (fixed !== xml) {
      files[path] = strToU8(fixed);
      changed = true;
    }
  }
  return changed ? zipSync(files) : bytes;
}

const text = (value: ExcelJS.CellValue) => String(plain(value) ?? "").trim();

/** Plain value of a cell (rich text, formulas, hyperlinks…). */
function plain(value: ExcelJS.CellValue): Cell {
  if (value === null || value === undefined) return null;
  if (value instanceof Date || typeof value !== "object") return value;
  if ("richText" in value) return value.richText.map((r) => r.text).join("");
  if ("result" in value) return plain(value.result as ExcelJS.CellValue);
  if ("text" in value) return String(value.text);
  if ("error" in value) return null;
  return String(value);
}

/** Rich text cell: native runs (bold, italic, underline, rgb colour) when styled, else a plain string. */
function richValue(html: string): ExcelJS.CellValue {
  const runs = htmlToRuns(html);
  if (!runs.length) return null;
  if (!isStyled(runs)) return runsToText(runs);
  return { richText: runs.map(toRichRun) };
}

const toRichRun = (run: Run): ExcelJS.RichText => ({
  text: run.text,
  font: {
    ...(run.bold && { bold: true }),
    ...(run.italic && { italic: true }),
    ...(run.underline && { underline: true }),
    ...(run.color && { color: { argb: `FF${run.color.slice(1).toUpperCase()}` } }),
  },
});

/** Rich HTML of a text cell: native runs or an escaped plain string. */
function richCell(value: ExcelJS.CellValue): Cell {
  if (value && typeof value === "object" && !(value instanceof Date) && "richText" in value)
    return runsToHtml(
      value.richText.map((r) => ({
        text: r.text,
        bold: r.font?.bold || undefined,
        italic: r.font?.italic || undefined,
        underline: !!r.font?.underline || undefined,
        color: /^[0-9A-F]{8}$/i.test(r.font?.color?.argb ?? "") ? `#${r.font!.color!.argb!.slice(2).toLowerCase()}` : undefined,
      })),
    );
  const text = plain(value);
  return text === null ? null : plainToHtml(String(text));
}

/** New empty workbook (for files created from the app). */
export async function emptyWorkbook(sheet: string) {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet(sheet);
  return new Uint8Array(await workbook.xlsx.writeBuffer());
}
