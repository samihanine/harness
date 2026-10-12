/**
 * Excel tables on SharePoint / OneDrive through the Graph workbook API. Rows are addressed by their
 * position in the table; the fields come from the hidden "_schema" sheet (or are guessed).
 */
import { graph } from "./ms";
import { ID_FIELD, SCHEMA_HEADER, coerce, fieldToRow, idPrefix, inferFields, isIdField, nextIds, rowToField, toCell, type Color, type Field, type FieldType, type Row } from "./schema";
import type { ExcelEntry } from "./store";

export type Table = { fields: Field[]; rows: Row[]; header: string[]; hasSchema: boolean };

export const shareId = (url: string) => "u!" + btoa(String.fromCharCode(...new TextEncoder().encode(url.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
const book = (e: Pick<ExcelEntry, "driveId" | "itemId">) => `/drives/${e.driveId}/items/${e.itemId}/workbook`;
const tablePath = (e: ExcelEntry) => `${book(e)}/tables/${encodeURIComponent(e.table)}`;
const SCHEMA_SHEET = "_schema";

/** Sharing link of an .xlsx → file reference, its server-relative URL (for Power BI) and its tables. */
export async function resolveExcel(link: string) {
  const item = await graph(`/shares/${shareId(link)}/driveItem?$select=id,name,parentReference`);
  const drive = await graph(`/drives/${item.parentReference.driveId}?$select=webUrl`);
  const folder = decodeURIComponent(item.parentReference.path?.split("root:")[1] ?? "");
  const ref = { driveId: item.parentReference.driveId as string, itemId: item.id as string };
  const { value: tables } = await graph(`${book(ref)}/tables?$select=name`);
  return { ...ref, name: item.name as string, fileUrl: `${drive.webUrl}${folder}/${item.name}`, tables: tables.map((t: { name: string }) => t.name) as string[] };
}

/**
 * Rows and fields. With a "_schema" sheet, only its fields are shown (other Excel columns are kept,
 * untouched, but hidden); without, every column is shown with a guessed type.
 */
export async function readTable(e: ExcelEntry): Promise<Table> {
  const [range, schema] = await Promise.all([
    graph(`${tablePath(e)}/range?$select=values,valueTypes,numberFormat`),
    graph(`${book(e)}/worksheets('${SCHEMA_SHEET}')/usedRange(valuesOnly=true)?$select=values`).catch(() => null),
  ]);
  const [header, ...data] = range.values as unknown[][];
  const names = header.map(String);
  const declared = ((schema?.values ?? []) as unknown[][]).slice(1).filter((r) => r[0]).map(rowToField);
  const fields = declared.length
    ? declared.filter((f) => names.includes(f.name))
    : inferFields(names, (range.valueTypes as string[][]).slice(1), (range.numberFormat as string[][]).slice(1), data);
  const rows = data.map((values) => {
    // Every column is kept (written back as is), typed when it is a field.
    const row: Row = Object.fromEntries(names.map((n, i) => [n, values[i]]));
    for (const f of fields)
      try {
        row[f.name] = coerce(f, row[f.name]);
      } catch {
        // kept as is (shown as invalid)
      }
    return row;
  });
  await fillIds(e, fields, rows);
  return { fields, rows, header: names, hasSchema: declared.length > 0 };
}

/** Rows without id get the next ones (existing ids never change), written to the id column. */
async function fillIds(e: ExcelEntry, fields: Field[], rows: Row[]) {
  const id = fields.find(isIdField);
  const empty = (v: unknown) => !String(toCell(v)).trim();
  // Blank rows (e.g. the one row of a new table) get none.
  const missing = id ? rows.filter((r) => empty(r[id.name]) && fields.some((f) => f !== id && f.type !== "boolean" && !empty(r[f.name]))) : [];
  if (!id || !missing.length) return;
  const ids = nextIds(idPrefix(e.table), rows.map((r) => r[id.name]), missing.length);
  missing.forEach((r, i) => (r[id.name] = ids[i]));
  await graph(`${tablePath(e)}/columns/${encodeURIComponent(id.name)}/dataBodyRange`, "PATCH", { values: rows.map((r) => [toCell(r[id.name])]) }).catch(() =>
    missing.forEach((r) => (r[id.name] = null)), // not saved: not shown either
  );
}

const check = (t: Table, values: Row) => {
  const out: Row = {};
  for (const [name, value] of Object.entries(values)) {
    const field = t.fields.find((f) => f.name === name);
    if (!field) throw new Error(`Unknown field "${name}" (fields: ${t.fields.map((f) => f.name).join(", ")})`);
    if (isIdField(field)) continue; // ids are given by the app only (addRows)
    out[name] = coerce(field, value);
  }
  return out;
};
/* Local versions of the writes, applied before the request (optimistic updates). */
export const withUpdates = (t: Table, changes: { index: number; values: Row }[]): Table => {
  const rows = [...t.rows];
  for (const { index, values } of changes) rows[index] = { ...rows[index], ...check(t, values) };
  return { ...t, rows };
};
export const withoutRows = (t: Table, indexes: number[]): Table => ({ ...t, rows: t.rows.filter((_, i) => !indexes.includes(i)) });
export const withRows = (t: Table, added: Row[]): Table => ({ ...t, rows: [...t.rows, ...added.map((r) => check(t, r))] });

const cells = (t: Table, row: Row) => t.header.map((name) => toCell(row[name]));

/** Sheet, first column and header line of a table (cached; columns are only added at the end). */
const anchors = new Map<string, Promise<{ sheet: string; left: number; top: number }>>();
function anchor(e: ExcelEntry) {
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
export async function updateRows(e: ExcelEntry, t: Table, changes: { index: number; values: Row }[]) {
  const { sheet, left, top } = await anchor(e);
  const ws = `${book(e)}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;
  for (const { index, values } of changes) {
    if (!t.rows[index]) throw new Error(`No row at index ${index}`);
    const checked = check(t, values);
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

export async function addRows(e: ExcelEntry, t: Table, rows: Row[]) {
  const required = t.fields.filter((f) => f.required);
  const checked = rows.map((r) => {
    const row = check(t, r);
    const missing = required.filter((f) => row[f.name] === undefined || row[f.name] === null || row[f.name] === "");
    if (missing.length) throw new Error(`Missing required field(s): ${missing.map(labelOrName).join(", ")}`);
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
}
const labelOrName = (f: Field) => f.label ?? f.name;

export async function deleteRows(e: ExcelEntry, indexes: number[]) {
  for (const index of [...indexes].sort((a, b) => b - a)) await graph(`${tablePath(e)}/rows/$/itemAt(index=${index})`, "DELETE");
}

/**
 * Writes the fields to the hidden "_schema" sheet, adds the Excel columns of new fields (columns are
 * never deleted: a removed field only stops being shown), then styles the sheet after the fields.
 */
export async function saveFields(e: ExcelEntry, t: Table, given: Field[], { style = true } = {}) {
  // Every table has an id column (first in the app, first Excel column when added).
  const fields = given.some(isIdField) ? given : [ID_FIELD, ...given];
  const names = new Set<string>();
  for (const f of fields) {
    if (!/^[\p{L}_][\p{L}\p{N}_ ]*$/u.test(f.name)) throw new Error(`Invalid field name "${f.name}"`);
    if (names.has(f.name)) throw new Error(`Duplicate field "${f.name}"`);
    names.add(f.name);
  }
  for (const f of fields.filter((f) => !t.header.includes(f.name)))
    await graph(`${tablePath(e)}/columns/add`, "POST", { index: isIdField(f) ? 0 : null, values: [[f.name], ...t.rows.map(() => [""])] });
  const sheet = `${book(e)}/worksheets('${SCHEMA_SHEET}')`;
  if (!(await graph(`${sheet}?$select=name`).then(() => true, () => false))) {
    await graph(`${book(e)}/worksheets/add`, "POST", { name: SCHEMA_SHEET });
    await graph(sheet, "PATCH", { visibility: "Hidden" });
  }
  await graph(`${sheet}/range(address='A1:G500')/clear`, "POST", { applyTo: "Contents" });
  const values = [[...SCHEMA_HEADER], ...fields.map(fieldToRow)];
  await graph(`${sheet}/range(address='A1:G${values.length}')`, "PATCH", { values });
  // Reordering only: the sheet styles do not change.
  if (style) await styleSheet(e, fields);
  return fields;
}

/** Option colors in the Excel file: light fill, dark text (same hues as the badges). */
const CELL_COLORS: Record<Color, { fill: string; font: string }> = {
  gray: { fill: "#F4F4F5", font: "#3F3F46" },
  red: { fill: "#FEE2E2", font: "#B91C1C" },
  orange: { fill: "#FFEDD5", font: "#C2410C" },
  amber: { fill: "#FEF3C7", font: "#92400E" },
  green: { fill: "#D1FAE5", font: "#047857" },
  teal: { fill: "#CCFBF1", font: "#0F766E" },
  blue: { fill: "#DBEAFE", font: "#1D4ED8" },
  violet: { fill: "#EDE9FE", font: "#6D28D9" },
  pink: { fill: "#FCE7F3", font: "#BE185D" },
};
/** Column width (points) by type: short for dates / yes-no, long and wrapped for long text. */
const WIDTH: Record<FieldType, number> = { boolean: 55, date: 85, integer: 70, number: 85, option: 110, string: 170, image: 220, text: 330 };
/** Short-text fields that hold long text (description, notes…): wide and wrapped like rich text. */
const isLongText = (f: Field) => f.type === "text" || (f.type === "string" && /desc|note|comment|summary|r[ée]sum|detail|content|remark/i.test(f.name));
const widthOf = (f: Field) => (isIdField(f) ? 75 : isLongText(f) ? WIDTH.text : WIDTH[f.type]);
/** Rows below the header that get the validation and formats (rows added later included). */
const STYLED_ROWS = 2000;

const letter = (n: number) => {
  let s = "";
  for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const columnIndex = (col: string) => [...col].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

/**
 * Look of the data sheet after the fields (best effort: a step refused by the API is skipped):
 * dark header, widths by type, wrapped rich text, dropdown + colors per option value.
 */
export async function styleSheet(e: ExcelEntry, fields: Field[]) {
  const step = (path: string, body: unknown, method = "PATCH") => graph(path, method, body).catch(() => undefined);
  const { address } = await graph(`${tablePath(e)}/range?$select=address`);
  const [, sheetName, firstCol, firstRow, lastCol] = address.match(/^'?(.*?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)/)!;
  const ws = `${book(e)}/worksheets('${encodeURIComponent(sheetName.replace(/'/g, "''"))}')`;
  const { values } = await graph(`${tablePath(e)}/headerRowRange?$select=values`);
  const header = (values[0] as unknown[]).map(String);
  const top = Number(firstRow);
  const left = columnIndex(firstCol);

  await step(`${tablePath(e)}`, { style: "TableStyleLight1", showBandedRows: false });
  await step(`${tablePath(e)}/headerRowRange/format`, { rowHeight: 26, verticalAlignment: "Center" });
  await step(`${tablePath(e)}/headerRowRange/format/fill`, { color: "#334155" });
  await step(`${tablePath(e)}/headerRowRange/format/font`, { color: "#FFFFFF", bold: true });
  await step(`${ws}/range(address='${firstCol}${top + 1}:${lastCol}${top + STYLED_ROWS}')/format`, { verticalAlignment: "Top" });

  for (const f of fields) {
    const i = header.indexOf(f.name);
    if (i < 0) continue;
    const col = letter(left + i);
    const cells = `${ws}/range(address='${col}${top + 1}:${col}${top + STYLED_ROWS}')`;
    const wrap = isLongText(f) || (f.type === "option" && !!f.multiple);
    await step(`${cells}/format`, { wrapText: wrap, horizontalAlignment: f.type === "boolean" ? "Center" : "General" });
    // Width set last, on a bounded range: through the whole-column address ('C:C') it was not
    // reliably applied, which left wrapped long text columns narrow.
    await step(`${ws}/range(address='${col}${top}:${col}${top + STYLED_ROWS}')/format`, { columnWidth: widthOf(f) });
    if (f.type === "date") await step(cells, { numberFormat: Array.from({ length: STYLED_ROWS }, () => ["yyyy-mm-dd"]) });
    // Validation and colors are rebuilt from the options each time.
    await step(`${cells}/dataValidation/clear`, {}, "POST");
    await step(`${cells}/conditionalFormats/clearAll`, {}, "POST");
    if (f.type !== "option" || !f.options?.length) continue;
    if (!f.multiple)
      await step(`${cells}/dataValidation`, {
        rule: { list: { inCellDropDown: true, source: f.options.map((o) => o.value).join(",") } },
        errorAlert: { showAlert: true, style: "Stop", title: "Invalid value", message: `Choose one of: ${f.options.map((o) => o.value).join(", ")}` },
      });
    for (const o of f.options) {
      const added = await graph(`${cells}/conditionalFormats/add`, "POST", { type: f.multiple ? "ContainsText" : "CellValue" }).catch(() => undefined);
      if (!added?.id) break; // conditional formats not available
      const format = `${cells}/conditionalFormats/${encodeURIComponent(added.id)}/${f.multiple ? "textComparison" : "cellValue"}`;
      const rule = f.multiple ? { operator: "Contains", text: o.value } : { operator: "EqualTo", formula1: `="${o.value.replace(/"/g, '""')}"` };
      // The format is a child resource: given inside the rule's PATCH it was ignored (rule without
      // fill, so transparent cells). Fill and font are set on their own paths.
      await step(format, { rule });
      await step(`${format}/format/fill`, { color: CELL_COLORS[o.color].fill });
      await step(`${format}/format/font`, { color: CELL_COLORS[o.color].font });
    }
  }
}

/* ---------------------------------- Images --------------------------------- */

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

/**
 * Uploads an image and returns its direct URL (stored in the cell): usable as is by Power BI
 * (image fields) and by the app. A sharing link for the organization is created so the report's
 * readers can see it.
 */
export async function uploadImage(e: ExcelEntry, file: File) {
  if (file.size > 4 * 1024 * 1024) throw new Error("Images up to 4 MB.");
  const target = e.imagesFolder
    ? { driveId: e.imagesFolder.driveId, parentId: e.imagesFolder.itemId, path: file.name }
    : { driveId: e.driveId, parentId: (await graph(`/drives/${e.driveId}/items/${e.itemId}?$select=parentReference`)).parentReference.id, path: `${e.name} images/${file.name}` };
  const base64 = await new Promise<string>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1]);
    reader.readAsDataURL(file);
  });
  const { uploadFile } = await import("@/server/ms");
  const item = await uploadFile({ data: { ...target, path: target.path.replace(/[^\w./ -]+/g, "-"), base64, type: file.type } });
  await graph(`/drives/${target.driveId}/items/${item.id}/createLink`, "POST", { type: "view", scope: "organization" }).catch(() => undefined);
  return (await graph(`/drives/${target.driveId}/items/${item.id}?$select=webUrl`)).webUrl as string;
}
