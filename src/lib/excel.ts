/**
 * Excel tables on SharePoint / OneDrive through the Graph workbook API. Rows are addressed by their
 * position in the table; the fields are the schema's (see schema.ts), also written in the workbook's
 * hidden "_schema" sheet when the schema is applied to the file.
 */
import { graph } from "./ms";
import { ID_FIELD, SCHEMA_HEADER, coerce, fieldToRow, idPrefix, inferFields, isIdField, nextIds, rowToField, toCell, type Color, type Field, type FieldType, type Row } from "./schema";
import type { ExcelRef } from "./store";
import { writeColumnRules } from "./xlsx-rules";

export type Table = { fields: Field[]; rows: Row[]; header: string[] };

export const shareId = (url: string) => "u!" + btoa(String.fromCharCode(...new TextEncoder().encode(url.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
const book = (e: Pick<ExcelRef, "driveId" | "itemId">) => `/drives/${e.driveId}/items/${e.itemId}/workbook`;
const tablePath = (e: ExcelRef) => `${book(e)}/tables/${encodeURIComponent(e.table)}`;
const SCHEMA_SHEET = "_schema";

/** Sharing link of an .xlsx → file reference, its server-relative URL (for Power BI) and its tables. */
export async function resolveExcel(link: string) {
  const item = await graph(`/shares/${shareId(link)}/driveItem?$select=id,name,parentReference,webUrl`);
  if (!/\.xlsx$/i.test(item.name)) throw new Error(`"${item.name}" is not an Excel file (.xlsx).`);
  const drive = await graph(`/drives/${item.parentReference.driveId}?$select=webUrl`);
  const folder = decodeURIComponent(item.parentReference.path?.split("root:")[1] ?? "");
  const ref = { driveId: item.parentReference.driveId as string, itemId: item.id as string };
  const { value: tables } = await graph(`${book(ref)}/tables?$select=name`);
  return { ...ref, name: (item.name as string).replace(/\.xlsx$/i, ""), webUrl: item.webUrl as string, fileUrl: `${drive.webUrl}${folder}/${item.name}`, tables: tables.map((t: { name: string }) => t.name) as string[] };
}

/** The columns of a table as fields: the workbook's "_schema" sheet when it has one, else guessed from the data. */
export async function detectFields(e: ExcelRef): Promise<Field[]> {
  const [range, schema] = await Promise.all([
    graph(`${tablePath(e)}/range?$select=values,valueTypes,numberFormat`),
    graph(`${book(e)}/worksheets('${SCHEMA_SHEET}')/usedRange(valuesOnly=true)?$select=values`).catch(() => null),
  ]);
  const [header, ...data] = range.values as unknown[][];
  const names = header.map(String);
  const declared = ((schema?.values ?? []) as unknown[][]).slice(1).filter((r) => r[0]).map(rowToField).filter((f) => names.includes(f.name));
  const guessed = inferFields(names, (range.valueTypes as string[][]).slice(1), (range.numberFormat as string[][]).slice(1), data);
  // Declared fields first (their order), then the columns the sheet does not describe.
  const fields = [...declared, ...guessed.filter((g) => !declared.some((d) => d.name === g.name))];
  return fields.some(isIdField) ? fields : [ID_FIELD, ...fields];
}

/**
 * Rows of the table, typed after the fields (other Excel columns are kept as they are, never shown).
 * Rows without id get one.
 */
export async function readRows(e: ExcelRef, fields: Field[]): Promise<Table> {
  const range = await graph(`${tablePath(e)}/range?$select=values`);
  const [header, ...data] = range.values as unknown[][];
  const names = header.map(String);
  const shown = fields.filter((f) => names.includes(f.name));
  const rows = data.map((values) => {
    const row: Row = Object.fromEntries(names.map((n, i) => [n, values[i]]));
    for (const f of shown)
      try {
        row[f.name] = coerce(f, row[f.name]);
      } catch {
        // kept as is (shown as invalid)
      }
    return row;
  });
  await fillIds(e, shown, rows);
  return { fields: shown, rows, header: names };
}

/** Rows without id get the next ones (existing ids never change), written to the id column. */
async function fillIds(e: ExcelRef, fields: Field[], rows: Row[]) {
  const id = fields.find(isIdField);
  const empty = (v: unknown) => !String(toCell(v)).trim();
  // Blank rows (e.g. the one row of a new table) get none.
  const missing = id ? rows.filter((r) => empty(r[id.name]) && fields.some((f) => f !== id && !empty(r[f.name]))) : [];
  if (!id || !missing.length) return;
  const ids = nextIds(idPrefix(e.table), rows.map((r) => r[id.name]), missing.length);
  missing.forEach((r, i) => (r[id.name] = ids[i]));
  await graph(`${tablePath(e)}/columns/${encodeURIComponent(id.name)}/dataBodyRange`, "PATCH", { values: rows.map((r) => [toCell(r[id.name])]) }).catch(() =>
    missing.forEach((r) => (r[id.name] = null)), // not saved: not shown either
  );
}

/** Why a row (by index) cannot take this id: empty, or already used by another row (case ignored). */
export function idError(t: Pick<Table, "fields" | "rows">, index: number, value: unknown) {
  const id = t.fields.find(isIdField);
  if (!id) return undefined;
  const v = String(value ?? "").trim();
  if (!v) return "The id cannot be empty.";
  const other = t.rows.findIndex((r, i) => i !== index && String(r[id.name] ?? "").trim().toLowerCase() === v.toLowerCase());
  return other >= 0 ? `The id "${v}" is already used by row ${other + 1}.` : undefined;
}

/** Typed values of a write. Ids are given by the app (addRows); they change only on an explicit edit (`index` given). */
const check = (t: Table, values: Row, index?: number) => {
  const out: Row = {};
  for (const [name, value] of Object.entries(values)) {
    const field = t.fields.find((f) => f.name === name);
    if (!field) throw new Error(`Unknown field "${name}" (fields: ${t.fields.map((f) => f.name).join(", ")})`);
    if (isIdField(field)) {
      if (index === undefined) continue;
      const error = idError(t, index, value);
      if (error) throw new Error(error);
      out[name] = String(value).trim();
      continue;
    }
    out[name] = coerce(field, value);
  }
  return out;
};
/* Local versions of the writes, applied before the request (optimistic updates). */
export const withUpdates = (t: Table, changes: { index: number; values: Row }[]): Table => {
  const next = { ...t, rows: [...t.rows] };
  // One after the other: two ids changed together are checked against each other.
  for (const { index, values } of changes) next.rows[index] = { ...next.rows[index], ...check(next, values, index) };
  return next;
};
export const withoutRows = (t: Table, indexes: number[]): Table => ({ ...t, rows: t.rows.filter((_, i) => !indexes.includes(i)) });
export const withRows = (t: Table, added: Row[]): Table => ({ ...t, rows: [...t.rows, ...added.map((r) => check(t, withDefaults(t.fields, r)))] });

const cells = (t: Table, row: Row) => t.header.map((name) => toCell(row[name]));

/** Sheet, first column and header line of a table (cached; columns are only added at the end). */
const anchors = new Map<string, Promise<{ sheet: string; left: number; top: number }>>();
function anchor(e: ExcelRef) {
  const key = tablePath(e);
  if (!anchors.has(key))
    anchors.set(
      key,
      graph(`${key}/range?$select=address`).then(({ address }) => {
        const [, sheet, col, row] = address.match(/^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)/);
        return { sheet: sheet.replace(/''/g, "'"), left: columnIndex(col), top: Number(row) };
      }),
    );
  return anchors.get(key)!.catch((error) => {
    anchors.delete(key);
    throw error;
  });
}

/**
 * Partial updates (`values` = changed fields only): one request per row, where unchanged cells are
 * `null` (left as they are: formulas and columns not shown by the app are never rewritten).
 */
export async function updateRows(e: ExcelRef, t: Table, changes: { index: number; values: Row }[]) {
  const { sheet, left, top } = await anchor(e);
  const ws = `${book(e)}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;
  for (const { index, values } of changes) {
    if (!t.rows[index]) throw new Error(`No row at index ${index}`);
    const checked = check(t, values, index);
    const unknown = Object.keys(checked).filter((name) => !t.header.includes(name));
    if (unknown.length) throw new Error(`Column(s) not in the Excel table: ${unknown.join(", ")}`);
    const line = top + 1 + index;
    const address = `${letter(left)}${line}:${letter(left + t.header.length - 1)}${line}`;
    const cellsRow = t.header.map((name) => (name in checked ? toCell(checked[name]) : null));
    await graph(`${ws}/range(address='${address}')`, "PATCH", { values: [cellsRow] }).catch((error) => {
      throw new Error(`Row ${index + 1} (${Object.keys(checked).join(", ")}): ${error instanceof Error ? error.message : error}`);
    });
    t.rows[index] = { ...t.rows[index], ...checked };
  }
}

/** A row with the fields' default values where it gives none. */
export const withDefaults = (fields: Field[], row: Row): Row => ({
  ...Object.fromEntries(fields.filter((f) => f.default !== undefined && !isIdField(f)).map((f) => [f.name, f.default])),
  ...Object.fromEntries(Object.entries(row).filter(([, v]) => v !== undefined)),
});

/**
 * Adds rows at the end (defaults filled, ids given). `partial`: a row started in the app, completed
 * afterwards (its required fields may still be empty).
 */
export async function addRows(e: ExcelRef, t: Table, rows: Row[], { partial = false } = {}) {
  const required = partial ? [] : t.fields.filter((f) => f.required && !isIdField(f));
  const checked = rows.map((r) => {
    const row = check(t, withDefaults(t.fields, r));
    const missing = required.filter((f) => row[f.name] === undefined || row[f.name] === null || row[f.name] === "" || (Array.isArray(row[f.name]) && !(row[f.name] as unknown[]).length));
    if (missing.length) throw new Error(`Missing required field(s): ${missing.map((f) => f.label ?? f.name).join(", ")}`);
    return row;
  });
  // Next ids from the file itself (not the screen): unique even when several adds are queued.
  const id = t.fields.find((f) => isIdField(f) && t.header.includes(f.name));
  if (id) {
    const { values } = await graph(`${tablePath(e)}/columns/${encodeURIComponent(id.name)}/dataBodyRange?$select=values`);
    const ids = nextIds(idPrefix(e.table), (values as unknown[][]).map((v) => v[0]), checked.length);
    checked.forEach((r, i) => (r[id.name] = ids[i]));
  }
  await graph(`${tablePath(e)}/rows/add`, "POST", { index: null, values: checked.map((r) => cells(t, r)) });
  return checked;
}

export async function deleteRows(e: ExcelRef, indexes: number[]) {
  for (const index of [...indexes].sort((a, b) => b - a)) await graph(`${tablePath(e)}/rows/$/itemAt(index=${index})`, "DELETE");
}

/* -------------------------------- Schema → file ------------------------------- */

/**
 * Writes a schema to its Excel file: missing columns added (columns are never deleted: a removed
 * field only stops being shown), the hidden "_schema" sheet rewritten, the rules written in the file
 * (choices' colors and dropdowns, ratings), then the look (header, grid, widths, row height). Only the
 * table's sheet and "_schema" change: the other sheets are left as they are. `progress` reports each
 * step (done / total). `appendOnly` (a table receiving a form's answers): new columns go at the end
 * only, so the columns the form writes to never move.
 */
export async function applySchemaToExcel(e: ExcelRef, fields: Field[], progress?: Progress, { appendOnly = false } = {}) {
  const total = 4 + fields.length;
  let done = 0;
  const step = (label: string) => progress?.(label, done++, total);
  step("Columns");
  const { values } = await graph(`${tablePath(e)}/headerRowRange?$select=values`);
  const header = (values[0] as unknown[]).map(String);
  const { value: rows } = await graph(`${tablePath(e)}/rows?$select=index`);
  for (const f of fields.filter((f) => !header.includes(f.name)))
    await graph(`${tablePath(e)}/columns/add`, "POST", { index: isIdField(f) && !appendOnly ? 0 : null, values: [[f.name], ...rows.map(() => [""])] });
  anchors.delete(tablePath(e));
  step("_schema sheet");
  const sheet = `${book(e)}/worksheets('${SCHEMA_SHEET}')`;
  if (!(await graph(`${sheet}?$select=name`).then(() => true, () => false))) {
    await graph(`${book(e)}/worksheets/add`, "POST", { name: SCHEMA_SHEET });
    await graph(sheet, "PATCH", { visibility: "Hidden" });
  }
  await graph(`${sheet}/range(address='A1:H500')/clear`, "POST", { applyTo: "Contents" });
  const table = [[...SCHEMA_HEADER], ...fields.map(fieldToRow)];
  await graph(`${sheet}/range(address='A1:H${table.length}')`, "PATCH", { values: table });
  step("Colors and dropdowns");
  // The file itself is changed (Graph has no rules): after the writes above have reached it.
  await new Promise((r) => setTimeout(r, 2500));
  const { address } = await graph(`${tablePath(e)}/range?$select=address`);
  const [, sheetName, firstCol, firstRow] = address.match(/^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)/)!;
  const now = ((await graph(`${tablePath(e)}/headerRowRange?$select=values`)).values[0] as unknown[]).map(String);
  const left = columnIndex(firstCol);
  const columns = fields.filter((f) => now.includes(f.name)).map((field) => ({ letter: letter(left + now.indexOf(field.name)), field }));
  const top = Number(firstRow);
  await writeColumnRules(e, sheetName.replace(/''/g, "'"), columns, top + 1, top + STYLED_ROWS, CELL_COLORS);
  // The file was replaced: the _schema sheet must still be there (written again if it was not saved yet).
  const kept = await graph(`${sheet}/usedRange(valuesOnly=true)?$select=rowCount`).then((r) => r.rowCount === table.length, () => false);
  if (!kept) {
    if (!(await graph(`${sheet}?$select=name`).then(() => true, () => false))) await graph(`${book(e)}/worksheets/add`, "POST", { name: SCHEMA_SHEET }).then(() => graph(sheet, "PATCH", { visibility: "Hidden" }));
    await graph(`${sheet}/range(address='A1:H${table.length}')`, "PATCH", { values: table });
  }
  await styleSheet(e, fields, (label) => step(label));
  progress?.("Done", total, total);
}

/** Reports a long task: what it does now, steps done / total. */
export type Progress = (label: string, done: number, total: number) => void;

/** Choice colors in the Excel file: the app's badges (10 % of the hue on white, dark text). */
const CELL_COLORS: Record<Color, { fill: string; font: string }> = {
  gray: { fill: "#F1F1F2", font: "#3F3F46" },
  red: { fill: "#FDECEC", font: "#B91C1C" },
  orange: { fill: "#FEF1E8", font: "#C2410C" },
  amber: { fill: "#FEF0DA", font: "#92400E" },
  green: { fill: "#E7F8F2", font: "#047857" },
  teal: { fill: "#E8F8F6", font: "#0F766E" },
  blue: { fill: "#EBF2FE", font: "#1D4ED8" },
  violet: { fill: "#F3EFFE", font: "#6D28D9" },
  pink: { fill: "#FDEDF5", font: "#BE185D" },
};
/** The look of every table: transparent cells, light gray header, dark gray 1 px grid. */
const LOOK = { header: "#F3F4F6", headerFont: "#111827", border: "#4B5563", rowHeight: 20, headerHeight: 24 };
/** Column width (points) by type: short for dates / ratings, long for long text. */
const WIDTH: Record<FieldType, number> = { date: 85, rating: 60, number: 85, option: 120, string: 170, url: 220, image: 120, file: 220, text: 300 };
const widthOf = (f: Field) => (isIdField(f) ? 80 : WIDTH[f.type]);
/** Rows below the header that get the validation and choice colors (rows added later included). */
const STYLED_ROWS = 2000;

const letter = (n: number) => {
  let s = "";
  for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const columnIndex = (col: string) => [...col].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
const EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"];

/**
 * Look of the data sheet after the fields, rewritten each time (a step refused by the API is skipped):
 * transparent cells, light gray bold header, dark gray 1 px borders around every cell, widths by
 * type, one row height (no wrapping), dates as dates. (Colors and dropdowns: xlsx-rules.ts.)
 */
async function styleSheet(e: ExcelRef, fields: Field[], progress?: (step: string) => void) {
  const step = (path: string, body: unknown, method = "PATCH") => graph(path, method, body).catch(() => undefined);
  const { address } = await graph(`${tablePath(e)}/range?$select=address`);
  const [, sheetName, firstCol, firstRow, lastCol, lastRow] = address.match(/^'?(.*?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/)!;
  const ws = `${book(e)}/worksheets('${encodeURIComponent(sheetName.replace(/''/g, "'").replace(/'/g, "''"))}')`;
  const { values } = await graph(`${tablePath(e)}/headerRowRange?$select=values`);
  const header = (values[0] as unknown[]).map(String);
  const top = Number(firstRow);
  const left = columnIndex(firstCol);
  const whole = `${ws}/range(address='${firstCol}${top}:${lastCol}${lastRow}')`;
  const body = `${ws}/range(address='${firstCol}${top + 1}:${lastCol}${Math.max(Number(lastRow), top + 1)}')`;

  progress?.("Look of the table");
  await step(`${tablePath(e)}`, { style: "TableStyleLight1", showBandedRows: false, showBandedColumns: false });
  await step(`${body}/format/fill/clear`, {}, "POST");
  await step(`${body}/format`, { rowHeight: LOOK.rowHeight, verticalAlignment: "Center", wrapText: false });
  await step(`${tablePath(e)}/headerRowRange/format`, { rowHeight: LOOK.headerHeight, verticalAlignment: "Center", wrapText: false });
  await step(`${tablePath(e)}/headerRowRange/format/fill`, { color: LOOK.header });
  await step(`${tablePath(e)}/headerRowRange/format/font`, { color: LOOK.headerFont, bold: true });
  for (const edge of EDGES) await step(`${whole}/format/borders/${edge}`, { style: "Continuous", color: LOOK.border, weight: "Thin" });

  for (const f of fields) {
    const i = header.indexOf(f.name);
    if (i < 0) {
      progress?.(`Column ${f.name}`);
      continue;
    }
    progress?.(`Column ${f.name}`);
    const col = letter(left + i);
    const cells = `${ws}/range(address='${col}${top + 1}:${col}${top + STYLED_ROWS}')`;
    // Width set on a bounded range: through the whole-column address ('C:C') it was not reliably applied.
    await step(`${ws}/range(address='${col}${top}:${col}${top + STYLED_ROWS}')/format`, { columnWidth: widthOf(f) });
    if (f.type === "date") await step(cells, { numberFormat: Array.from({ length: STYLED_ROWS }, () => ["yyyy-mm-dd"]) });
  }
}

/** A table on the first sheet of an empty workbook, with the given header. */
export async function createTable(file: Pick<ExcelRef, "driveId" | "itemId">, header: string[]) {
  const { value } = await graph(`${book(file)}/worksheets?$select=name`);
  const sheet = `${book(file)}/worksheets('${encodeURIComponent(value[0].name)}')`;
  const address = `A1:${letter(header.length - 1)}1`;
  await graph(`${sheet}/range(address='${address}')`, "PATCH", { values: [header] });
  return (await graph(`${sheet}/tables/add`, "POST", { address, hasHeaders: true })).name as string;
}

/* --------------------------------- Preview --------------------------------- */

/**
 * URL showing the file in Excel for the web, read only, in an iframe (Graph "preview": the access
 * token is in the URL, valid a few minutes for opening it; no browser session needed).
 */
export async function previewUrl(e: Pick<ExcelRef, "driveId" | "itemId">) {
  const { getUrl } = await graph(`/drives/${e.driveId}/items/${e.itemId}/preview`, "POST", {});
  return getUrl as string;
}

/* ---------------------------------- Files ---------------------------------- */

const urls = new Map<string, Promise<string | null>>();

/** Displayable URL of an image cell: SharePoint / OneDrive links are read server side with the user's rights. */
export function imageUrl(value: unknown): Promise<string | null> {
  const url = String(value ?? "").trim();
  if (!/^https?:\/\//i.test(url)) return Promise.resolve(null);
  if (!/sharepoint\.com|1drv\.ms|onedrive\.live\.com/i.test(url)) return Promise.resolve(url);
  if (!urls.has(url)) urls.set(url, import("@/server/ms").then(({ downloadFile }) => downloadFile({ data: { url } })).catch(() => null));
  return urls.get(url)!;
}

/** Sharing link of a SharePoint / OneDrive folder → reference. */
export async function resolveFolder(link: string) {
  const item = await graph(`/shares/${shareId(link)}/driveItem?$select=id,name,folder,parentReference`);
  if (!item.folder) throw new Error("This link is not a folder.");
  return { link, driveId: item.parentReference.driveId as string, itemId: item.id as string, name: item.name as string };
}

export const toBase64 = (file: Blob) =>
  new Promise<string>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1]);
    reader.readAsDataURL(file);
  });

/**
 * Uploads a file into a folder and returns its direct URL (stored in the cell): usable as is by Power
 * BI (image fields) and by the app. A sharing link for the organization is created so readers can see it.
 */
export async function uploadFileTo(folder: { driveId: string; itemId: string }, file: File) {
  if (file.size > 4 * 1024 * 1024) throw new Error("Files up to 4 MB.");
  const { uploadFile } = await import("@/server/ms");
  const item = await uploadFile({ data: { driveId: folder.driveId, parentId: folder.itemId, path: file.name.replace(/[^\w./ -]+/g, "-"), base64: await toBase64(file), type: file.type } });
  await graph(`/drives/${folder.driveId}/items/${item.id}/createLink`, "POST", { type: "view", scope: "organization" }).catch(() => undefined);
  return (await graph(`/drives/${folder.driveId}/items/${item.id}?$select=webUrl`)).webUrl as string;
}
