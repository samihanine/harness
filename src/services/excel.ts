/**
 * Excel tables on SharePoint / OneDrive through the Graph workbook API. Rows are addressed by their
 * position in the table; the columns are the table's own, read from its header each time (the app
 * never types, adds or moves them: no gap between the app and the file).
 */
import { graph, request, uploadFile } from "./ms";
import { coerce, idPrefix, isIdField, nextIds, toCell, type Field, type Row } from "@/features/excel/schema";
import type { ExcelRef } from "@/lib/store";

export type Table = { fields: Field[]; rows: Row[]; header: string[] };

export const shareId = (url: string) => "u!" + btoa(String.fromCharCode(...new TextEncoder().encode(url.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
const book = (e: Pick<ExcelRef, "driveId" | "itemId">) => `/drives/${e.driveId}/items/${e.itemId}/workbook`;
const tablePath = (e: ExcelRef) => `${book(e)}/tables/${encodeURIComponent(e.table)}`;

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

/** The columns of a table, from its header row (a text field each; the id column, when there is one, kept as such). */
const headerFields = (names: string[]): Field[] => names.map((name) => ({ name, type: isIdField({ name }) ? "string" : "text" }));

export async function detectFields(e: ExcelRef): Promise<Field[]> {
  const { values } = await graph(`${tablePath(e)}/headerRowRange?$select=values`);
  return headerFields((values[0] as unknown[]).map(String));
}

/** Rows of the table and its columns, as the file has them now. Rows without id get one (when the table has an id column). */
export async function readRows(e: ExcelRef): Promise<Table> {
  const range = await graph(`${tablePath(e)}/range?$select=values`);
  const [header, ...data] = range.values as unknown[][];
  const names = header.map(String);
  const fields = headerFields(names);
  const rows = data.map((values) => Object.fromEntries(names.map((n, i) => [n, values[i]])) as Row);
  await fillIds(e, fields, rows);
  return { fields, rows, header: names };
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

/** A field by its name, the case and spaces aside ("Expected Result" is the column "Expected result"). */
export const fieldOf = (t: Pick<Table, "fields">, name: string) => {
  const key = (s: string) => s.toLowerCase().replace(/[\s_]+/g, " ").trim();
  return t.fields.find((f) => f.name === name) ?? t.fields.find((f) => key(f.name) === key(name));
};

/** Typed values of a write. Ids are given by the app (addRows); they change only on an explicit edit (`index` given). */
const check = (t: Table, values: Row, index?: number) => {
  const out: Row = {};
  for (const [name, value] of Object.entries(values)) {
    const field = fieldOf(t, name);
    if (!field) throw new Error(`Unknown field "${name}" (fields: ${t.fields.map((f) => f.name).join(", ")})`);
    if (isIdField(field)) {
      if (index === undefined) continue;
      const error = idError(t, index, value);
      if (error) throw new Error(error);
      out[name] = String(value).trim();
      continue;
    }
    out[field.name] = coerce(field, value);
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

/** Reports a long task: what it does now, steps done / total. */
export type Progress = (label: string, done: number, total: number) => void;

/** The look of every table: light gray bold header, gray 1 px grid, wrapped text. */
const LOOK = { header: "#E5E7EB", headerFont: "#111827", border: "#9CA3AF", headerHeight: 24, idWidth: 80, width: 180 };
const EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"];
const letter = (n: number) => {
  let s = "";
  for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const columnIndex = (col: string) => [...col].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

/**
 * "Sync the formatting": only the look of the table changes, never its columns or values — gray bold
 * header, borders around every cell, column widths, text wrapped and rows fitted to it (a step the
 * API refuses is skipped).
 */
export async function formatTable(e: ExcelRef) {
  const step = (path: string, body: unknown, method = "PATCH") => graph(path, method, body).catch(() => undefined);
  const { address } = await graph(`${tablePath(e)}/range?$select=address`);
  const [, sheetName, firstCol, firstRow, , lastRow] = address.match(/^'?(.*?)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/)!;
  const ws = `${book(e)}/worksheets('${encodeURIComponent(sheetName.replace(/''/g, "'").replace(/'/g, "''"))}')`;
  const header = ((await graph(`${tablePath(e)}/headerRowRange?$select=values`)).values[0] as unknown[]).map(String);
  const top = Number(firstRow);
  const left = columnIndex(firstCol);
  await step(tablePath(e), { style: "TableStyleLight1", showBandedRows: false, showBandedColumns: false });
  await step(`${tablePath(e)}/dataBodyRange/format/fill/clear`, {}, "POST");
  await step(`${tablePath(e)}/headerRowRange/format`, { rowHeight: LOOK.headerHeight, verticalAlignment: "Center", wrapText: false });
  await step(`${tablePath(e)}/headerRowRange/format/fill`, { color: LOOK.header });
  await step(`${tablePath(e)}/headerRowRange/format/font`, { color: LOOK.headerFont, bold: true });
  for (const edge of EDGES) await step(`${tablePath(e)}/range/format/borders/${edge}`, { style: "Continuous", color: LOOK.border, weight: "Thin" });
  // Widths on a bounded range of each column: through the whole-column address ('C:C') they were not reliably applied.
  for (const [i, name] of header.entries()) {
    const col = letter(left + i);
    await step(`${ws}/range(address='${col}${top}:${col}${Math.max(Number(lastRow), top + 1)}')/format`, { columnWidth: isIdField({ name }) ? LOOK.idWidth : LOOK.width });
  }
  await step(`${tablePath(e)}/dataBodyRange/format`, { wrapText: true, verticalAlignment: "Top" });
  await step(`${tablePath(e)}/dataBodyRange/format/autofitRows`, {}, "POST");
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

/** Displayable URL of an image cell: SharePoint / OneDrive links are read through Graph with the user's rights. */
export function imageUrl(value: unknown): Promise<string | null> {
  const url = String(value ?? "").trim();
  if (!/^https?:\/\//i.test(url)) return Promise.resolve(null);
  if (!/sharepoint\.com|1drv\.ms|onedrive\.live\.com/i.test(url)) return Promise.resolve(url);
  if (!urls.has(url))
    urls.set(
      url,
      request("graph", `/shares/${shareId(url)}/driveItem/content`)
        .then((r) => (r.ok ? r.blob() : null))
        .then((b) => (b ? URL.createObjectURL(b) : null))
        .catch(() => null),
    );
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
  const item = await uploadFile({ driveId: folder.driveId, parentId: folder.itemId, path: file.name.replace(/[^\w./ -]+/g, "-"), file });
  await graph(`/drives/${folder.driveId}/items/${item.id}/createLink`, "POST", { type: "view", scope: "organization" }).catch(() => undefined);
  return (await graph(`/drives/${folder.driveId}/items/${item.id}?$select=webUrl`)).webUrl as string;
}
