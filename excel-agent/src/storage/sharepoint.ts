/**
 * Excel file on SharePoint / OneDrive, edited in place with the Graph Excel API (no download):
 * one persistent workbook session, the rows kept as an Excel table, changes queued then sent
 * in a single JSON batch (one row = one range write; appends grouped in one call).
 */
import { htmlToRuns, plainToHtml, runsToText } from "@/lib/rich";
import type { BatchRequest } from "./graph";
import { GraphError, batch, graph } from "./graph";
import type { Cell, ColumnSpec, DriveRef, TableBackend, Values } from "./types";
import { imageFormula, rowHeight } from "./types";

type RangeData = { address: string; values: Cell[][] };

const column = (n: number) => {
  let s = "";
  for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const same = (a: Cell, b: Cell) => String(a ?? "") === String(b ?? "");

/** Plain table: the look comes from the cell formats set by format(); also marks a formatted table. */
const TABLE_STYLE = "TableStyleLight1";
/** Rows covered by row heights, borders, dropdowns and option colours. */
const FORMATTED_ROWS = 5000;

export class SharePointTable implements TableBackend {
  readonly pollEvery = 15_000;
  private session?: Promise<string>;
  private sheet = "";
  private table = "";
  private header: string[] = [];
  /** Sheet position of the table header (row number, first column index). */
  private top = 1;
  private left = 0;
  /** Data rows in table order (null = empty row, reused by appends). */
  private data: (Values | null)[] = [];
  private queue: BatchRequest[] = [];
  private dates = new Set<string>();
  /** Rich text columns: Graph writes values only, so they are stored as plain text. */
  private rich = new Set<string>();
  /** Computed columns (IMAGE of a link column): written as formulas, never read as data. */
  private images = new Map<string, string>();

  constructor(private file: DriveRef) {}

  private get base() {
    return `/drives/${this.file.driveId}/items/${this.file.itemId}/workbook`;
  }

  /** Persistent workbook session: changes are saved, and calls are much faster than sessionless ones. */
  private sessionId() {
    this.session ??= graph<{ id: string }>(`${this.base}/createSession`, { method: "POST", json: { persistChanges: true } })
      .then((s) => s.id)
      .catch((error) => {
        this.session = undefined;
        throw error;
      });
    return this.session;
  }

  /** Workbook call with the session; an expired session is recreated once. */
  private async call<T>(path: string, init: RequestInit & { json?: unknown } = {}, retry = true): Promise<T> {
    try {
      return await graph<T>(`${this.base}${path}`, { ...init, headers: { "workbook-session-id": await this.sessionId(), ...init.headers } });
    } catch (error) {
      if (retry && error instanceof GraphError && /session/i.test(error.code + error.message)) {
        this.session = undefined;
        return this.call(path, init, false);
      }
      throw error;
    }
  }

  async open(sheet: string, specs: ColumnSpec[]) {
    this.sheet = sheet;
    this.dates = new Set(specs.filter((s) => s.date).map((s) => s.name));
    this.rich = new Set(specs.filter((s) => s.rich).map((s) => s.name));
    this.images = new Map(specs.filter((s) => s.imageOf).map((s) => [s.name, s.imageOf!]));
    const ws = `/worksheets/${encodeURIComponent(sheet)}`;
    await this.call(`${ws}?$select=name`).catch(async (error) => {
      if (!(error instanceof GraphError) || error.status !== 404) throw error;
      await this.call("/worksheets/add", { method: "POST", json: { name: sheet } });
    });

    // The rows live in an Excel table: the sheet's first one, or one created on the header.
    const { value: tables } = await this.call<{ value: { name: string; style?: string }[] }>(`${ws}/tables?$select=name,style`);
    let created = false;
    const styled = tables[0]?.style === TABLE_STYLE;
    if (tables[0]) this.table = tables[0].name;
    else {
      const used = await this.call<RangeData>(`${ws}/usedRange(valuesOnly=true)?$select=address,values`);
      const empty = used.values.every((row) => row.every((v) => v === "" || v === null));
      let address = used.address.split("!")[1];
      if (empty) {
        address = `A1:${column(specs.length - 1)}1`;
        await this.call(`${ws}/range(address='${address}')`, { method: "PATCH", json: { values: [specs.map((s) => s.name)] } });
      }
      const table = await this.call<{ name: string }>(`${ws}/tables/add`, { method: "POST", json: { address, hasHeaders: true } });
      this.table = table.name;
      created = true;
    }

    let range = await this.readTable();
    const missing = specs.filter((s) => !this.header.includes(s.name));
    for (const spec of missing)
      await this.call(`/tables/${encodeURIComponent(this.table)}/columns/add`, {
        method: "POST",
        json: { index: null, values: [[spec.name], ...range.values.slice(1).map(() => [""])] },
      });
    if (missing.length) range = await this.readTable();
    if (created || missing.length || !styled) await this.format(specs);
    // New picture columns: their formula on the existing rows.
    if (missing.some((s) => s.imageOf)) this.data.forEach((row, i) => row && this.queue.push(this.rowRequest(i, row)));

    // Rows without id get one.
    this.data.forEach((row, i) => {
      if (row && !row.id) this.update([{ id: "", values: { id: crypto.randomUUID() } }], i);
    });
    await this.flush();
    return this.data.filter((r): r is Values => !!r);
  }

  private async readTable() {
    const range = await this.call<RangeData>(`/tables/${encodeURIComponent(this.table)}/range?$select=address,values`);
    const start = /!\$?([A-Z]+)\$?(\d+)/.exec(range.address)!;
    this.left = [...start[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
    this.top = Number(start[2]);
    this.header = range.values[0].map((v) => String(v ?? "").trim());
    this.data = range.values.slice(1).map((row) => (row.every((v) => v === "" || v === null) ? null : this.record(row)));
    return range;
  }

  /** Row values by column; Excel date serials of date columns become "YYYY-MM-DD". */
  private record(row: Cell[]): Values {
    return Object.fromEntries(
      this.header.flatMap((h, i): [string, Cell][] => {
        if (this.images.has(h)) return [];
        const v = row[i] === "" ? null : row[i];
        if (this.dates.has(h) && typeof v === "number") return [[h, new Date(Math.round((v - 25569) * 86_400_000)).toISOString().slice(0, 10)]];
        if (this.rich.has(h) && v !== null) return [[h, plainToHtml(String(v))]];
        return [[h, v]];
      }),
    );
  }

  /**
   * Look of the sheet (best effort: each step is skipped if the tenant or API refuses it): plain
   * table, dark header, taller rows, thin borders, dropdowns with an error alert on option
   * columns, option cells coloured by value.
   */
  private async format(specs: ColumnSpec[]) {
    const ws = `/worksheets/${encodeURIComponent(this.sheet)}`;
    const table = `/tables/${encodeURIComponent(this.table)}`;
    const last = column(this.left + this.header.length - 1);
    const all = `${column(this.left)}${this.top}:${last}${this.top + FORMATTED_ROWS}`;
    const body = `${column(this.left)}${this.top + 1}:${last}${this.top + FORMATTED_ROWS}`;
    const step = (path: string, json: unknown, method = "PATCH") => this.call(path, { method, json }).catch(() => undefined);

    await step(table, { style: TABLE_STYLE, showBandedRows: false, showBandedColumns: false });
    await step(`${table}/headerRowRange/format`, { rowHeight: 28, verticalAlignment: "Center" });
    await step(`${table}/headerRowRange/format/fill`, { color: "#334155" });
    await step(`${table}/headerRowRange/format/font`, { color: "#FFFFFF", bold: true });
    await step(`${ws}/range(address='${body}')/format`, { rowHeight: rowHeight(specs), verticalAlignment: "Center" });
    for (const side of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"])
      await step(`${ws}/range(address='${all}')/format/borders/${side}`, { style: "Continuous", weight: "Thin", color: "#D4D4D8" });

    for (const spec of specs) {
      const letter = column(this.left + this.header.indexOf(spec.name));
      const cells = `${ws}/range(address='${letter}${this.top + 1}:${letter}${this.top + FORMATTED_ROWS}')`;
      if (spec.rich || spec.wrap) await step(`${cells}/format`, { wrapText: true, verticalAlignment: "Top" });
      if (spec.imageOf) await step(`${cells}/format`, { horizontalAlignment: "Center" });
      if (spec.list?.length)
        await step(`${cells}/dataValidation`, {
          rule: { list: { inCellDropDown: true, source: spec.list.join(",") } },
          errorAlert: { showAlert: true, style: "Stop", title: "Invalid value", message: `Choose one of: ${spec.list.join(", ")}` },
        });
      for (const { value, fill, font } of spec.colors ?? []) {
        const kind = spec.multiple ? "textComparison" : "cellValue";
        const added = await this.call<{ id: string }>(`${cells}/conditionalFormats/add`, { method: "POST", json: { type: spec.multiple ? "ContainsText" : "CellValue" } }).catch(() => undefined);
        if (!added) break; // conditional formats not available here
        const format = { fill: { color: fill }, font: { color: font } };
        const rule = spec.multiple ? { operator: "Contains", text: value } : { operator: "EqualTo", formula1: `="${value.replace(/"/g, '""')}"` };
        await step(`${cells}/conditionalFormats/${encodeURIComponent(added.id)}/${kind}`, { format, rule });
      }
    }
  }

  private indexOf(id: string) {
    const index = this.data.findIndex((r) => r?.id === id);
    if (index < 0) throw new Error(`No row with id ${id}`);
    return index;
  }

  /** Value written to Graph: null would mean "unchanged", so empty cells are "". */
  private cell(column: string, value: Cell | undefined) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    if (this.rich.has(column) && typeof value === "string") return runsToText(htmlToRuns(value));
    return value;
  }

  private rowRequest(index: number, row: Values): BatchRequest {
    const line = this.top + 1 + index;
    const address = `${column(this.left)}${line}:${column(this.left + this.header.length - 1)}${line}`;
    // "formulas" takes constants and formulas: the picture columns get their IMAGE() of the link.
    const formulas = [
      this.header.map((h) => {
        const source = this.images.get(h);
        if (!source) return this.cell(h, row[h]);
        const link = `${column(this.left + this.header.indexOf(source))}${line}`;
        return `=${imageFormula(link, false)}`;
      }),
    ];
    return { method: "PATCH", url: `${this.base}/worksheets/${encodeURIComponent(this.sheet)}/range(address='${address}')`, body: { formulas } };
  }

  update(changes: { id: string; values: Values }[], index?: number) {
    for (const { id, values } of changes) {
      const i = index ?? this.indexOf(id);
      const row = { ...this.data[i], ...values };
      if (Object.keys(values).every((k) => same(this.data[i]?.[k] ?? null, values[k]))) continue;
      this.data[i] = row;
      this.queue.push(this.rowRequest(i, row));
    }
  }

  append(rows: Values[]) {
    const fresh: Values[] = [];
    for (const row of rows) {
      const empty = this.data.indexOf(null);
      if (empty >= 0) {
        this.data[empty] = row;
        this.queue.push(this.rowRequest(empty, row));
      } else fresh.push(row);
    }
    if (!fresh.length) return;
    const start = this.data.length;
    this.data.push(...fresh);
    const values = fresh.map((row) => this.header.map((h) => (this.images.has(h) ? "" : this.cell(h, row[h]))));
    this.queue.push({ method: "POST", url: `${this.base}/tables/${encodeURIComponent(this.table)}/rows/add`, body: { index: null, values } });
    // Rows are added without formulas, then the picture formulas are set.
    if (this.images.size) fresh.forEach((row, k) => this.queue.push(this.rowRequest(start + k, row)));
  }

  remove(ids: string[]) {
    // From the bottom so the remaining indexes stay valid.
    for (const index of ids.map((id) => this.indexOf(id)).sort((a, b) => b - a)) {
      this.data.splice(index, 1);
      this.queue.push({ method: "DELETE", url: `${this.base}/tables/${encodeURIComponent(this.table)}/rows/$/itemAt(index=${index})` });
    }
  }

  async flush() {
    if (!this.queue.length) return;
    const requests = this.queue.splice(0);
    const session = await this.sessionId();
    try {
      await batch(requests.map((r) => ({ ...r, headers: { "workbook-session-id": session } })));
      // Changes of an open session are only committed to the file (and visible in Excel) when it closes.
      this.session = undefined;
      await graph(`${this.base}/closeSession`, { method: "POST", headers: { "workbook-session-id": session } }).catch(() => undefined);
    } catch (error) {
      // The state here no longer matches the file: reload it on the next check.
      this.data = [];
      throw error;
    }
  }

  async changedOutside() {
    if (this.queue.length) return false;
    const range = await this.call<RangeData>(`/tables/${encodeURIComponent(this.table)}/range?$select=values`);
    const rows = range.values.slice(1);
    if (rows.length !== this.data.length) return true;
    return rows.some((row, i) => {
      const record = this.record(row);
      return this.header.some((h) => !this.images.has(h) && !same(record[h], this.data[i]?.[h] ?? null));
    });
  }
}

/** Creates an .xlsx in a SharePoint / OneDrive folder. */
export async function createDriveFile(folder: DriveRef, name: string, bytes: Uint8Array): Promise<DriveRef> {
  const file = name.toLowerCase().endsWith(".xlsx") ? name : `${name}.xlsx`;
  const item = await graph<{ id: string; name: string; webUrl: string; parentReference: { driveId: string } }>(
    `/drives/${folder.driveId}/items/${folder.itemId}:/${encodeURIComponent(file)}:/content?@microsoft.graph.conflictBehavior=fail`,
    { method: "PUT", body: bytes as BodyInit, headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } },
  );
  return { driveId: item.parentReference.driveId, itemId: item.id, name: item.name, webUrl: item.webUrl };
}
