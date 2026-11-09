/**
 * Excel tables on SharePoint / OneDrive through the Graph workbook API: generic reads and writes of a
 * table (header, rows by position, columns), many requests sent together ($batch), the file's
 * version (to see changes made in Excel), a hidden sheet for the app's own data, and the look of the
 * table. What the API cannot do (validation lists, conditional colors) is written into the file
 * itself by features/excel/xlsx.ts.
 */
import { call, graph } from "./ms";

/** An Excel table on SharePoint / OneDrive, as found from its link. */
export type ExcelRef = { link: string; driveId: string; itemId: string; name: string; table: string; fileUrl: string; webUrl?: string };
export type Cell = string | number | boolean | null;

export const shareId = (url: string) => "u!" + btoa(String.fromCharCode(...new TextEncoder().encode(url.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
export const book = (e: Pick<ExcelRef, "driveId" | "itemId">) => `/drives/${e.driveId}/items/${e.itemId}/workbook`;
export const tablePath = (e: ExcelRef) => `${book(e)}/tables/${encodeURIComponent(e.table)}`;
const sheetPath = (e: Pick<ExcelRef, "driveId" | "itemId">, sheet: string) => `${book(e)}/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;

export const letter = (n: number) => {
  let s = "";
  for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
export const columnIndex = (col: string) => [...col].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

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

/** The file's version: changes when anyone saves it (Excel, the app, another tool). */
export async function versionOf(e: Pick<ExcelRef, "driveId" | "itemId">) {
  const m = await graph(`/drives/${e.driveId}/items/${e.itemId}?$select=eTag,lastModifiedDateTime,lastModifiedBy`);
  return { eTag: m.eTag as string, at: m.lastModifiedDateTime as string, by: (m.lastModifiedBy?.user?.displayName as string | undefined) ?? undefined };
}

/** Where the table is: its sheet, first column, header line and size. */
export type Anchor = { sheet: string; left: number; top: number; width: number; height: number };
const anchorOf = (address: string): Anchor => {
  const [, sheet, c1, r1, c2, r2] = address.match(/^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?/)!;
  const left = columnIndex(c1);
  const top = Number(r1);
  return { sheet: sheet.replace(/''/g, "'"), left, top, width: columnIndex(c2 ?? c1) - left + 1, height: Number(r2 ?? r1) - top + 1 };
};

/** The table as the file has it now: its header, its rows (cells by position) and where it is. */
export async function readTable(e: ExcelRef): Promise<{ header: string[]; rows: Cell[][]; anchor: Anchor }> {
  const range = await graph(`${tablePath(e)}/range?$select=values,address`);
  const [header, ...rows] = range.values as Cell[][];
  return { header: header.map((h) => String(h ?? "")), rows, anchor: anchorOf(range.address) };
}

/* ----------------------------------- Batch ----------------------------------- */

export type BatchRequest = { method: "GET" | "POST" | "PATCH" | "DELETE"; url: string; body?: unknown };

/**
 * Requests sent together (Graph $batch: 20 per call), in order: each waits for the one before (row
 * positions stay right). The first failure stops the rest and is thrown with its message.
 */
export async function batch(requests: BatchRequest[]) {
  const results: unknown[] = [];
  for (let start = 0; start < requests.length; start += 20) {
    const chunk = requests.slice(start, start + 20).map((r, i) => ({
      id: String(i + 1),
      method: r.method,
      url: r.url,
      ...(i > 0 && { dependsOn: [String(i)] }),
      ...(r.body !== undefined && { body: r.body, headers: { "Content-Type": "application/json" } }),
    }));
    const { responses } = await call<{ responses: { id: string; status: number; body?: any }[] }>("graph", "/$batch", "POST", { requests: chunk });
    for (const r of [...responses].sort((a, b) => Number(a.id) - Number(b.id))) {
      if (r.status >= 400) throw new Error(r.body?.error?.message ?? `Excel refused a change (${r.status}).`);
      results.push(r.body);
    }
  }
  return results;
}

/* ------------------------------------ Rows ----------------------------------- */

/** Writes whole rows (cells in the file's column order) at their positions in the table's body. */
export const writeRows = (e: ExcelRef, rows: { index: number; cells: Cell[] }[]) =>
  batch(rows.map(({ index, cells }) => ({ method: "PATCH", url: `${tablePath(e)}/rows/itemAt(index=${index})`, body: { values: [cells] } })));

/** Adds rows at the end of the table. */
export const appendRows = (e: ExcelRef, rows: Cell[][]) => (rows.length ? batch([{ method: "POST", url: `${tablePath(e)}/rows/add`, body: { index: null, values: rows } }]) : Promise.resolve([]));

/** Deletes rows by position (from the last: the others keep theirs). */
export const deleteRows = (e: ExcelRef, indexes: number[]) => batch([...new Set(indexes)].sort((a, b) => b - a).map((i) => ({ method: "DELETE", url: `${tablePath(e)}/rows/itemAt(index=${i})` })));

/* ---------------------------------- Columns ---------------------------------- */

/** Adds columns at the end of the table (their header names). */
export const addColumns = (e: ExcelRef, names: string[]) => batch(names.map((name) => ({ method: "POST", url: `${tablePath(e)}/columns/add`, body: { index: null, name } })));

export const deleteColumns = (e: ExcelRef, names: string[]) => batch(names.map((name) => ({ method: "DELETE", url: `${tablePath(e)}/columns/${encodeURIComponent(name)}` })));

/** Rewrites the whole table (header and rows) in place: columns reordered or renamed, same width. */
export async function rewriteTable(e: ExcelRef, header: string[], rows: Cell[][]) {
  const { anchor } = await readTable(e);
  if (header.length !== anchor.width) throw new Error("The table's width changed meanwhile: read it again.");
  const address = `${letter(anchor.left)}${anchor.top}:${letter(anchor.left + anchor.width - 1)}${anchor.top + rows.length}`;
  await graph(`${sheetPath(e, anchor.sheet)}/range(address='${address}')`, "PATCH", { values: [header, ...rows] });
}

/* ------------------------------- The app's sheet ------------------------------ */

/** A hidden sheet keeping the app's data with the file (the columns' types and colors, the views): a JSON text in column A. */
const APP_SHEET = "_hub";
const CHUNK = 30_000;

export async function readAppData<T>(e: ExcelRef): Promise<T | undefined> {
  try {
    const { values } = await graph(`${sheetPath(e, APP_SHEET)}/range(address='A1:A40')?$select=values`);
    const text = (values as Cell[][]).map((r) => String(r[0] ?? "")).join("");
    return text ? (JSON.parse(text) as T) : undefined;
  } catch {
    return undefined;
  }
}

export async function writeAppData(e: ExcelRef, data: unknown) {
  const text = JSON.stringify(data);
  const chunks = Array.from({ length: 40 }, (_, i) => [text.slice(i * CHUNK, (i + 1) * CHUNK)]);
  if (text.length > 40 * CHUNK) throw new Error("The app's data is too large for its sheet.");
  const sheets = (await graph(`${book(e)}/worksheets?$select=name`)).value as { name: string }[];
  if (!sheets.some((s) => s.name === APP_SHEET)) await graph(`${book(e)}/worksheets/add`, "POST", { name: APP_SHEET });
  await batch([
    { method: "PATCH", url: `${sheetPath(e, APP_SHEET)}/range(address='A1:A40')`, body: { values: chunks } },
    { method: "PATCH", url: sheetPath(e, APP_SHEET), body: { visibility: "Hidden" } },
  ]);
}

/* ---------------------------------- The look ---------------------------------- */

export type ColumnLook = { width: number; wrap: boolean; numberFormat?: string };

/**
 * The look of the table: light gray bold header, thin gray grid, each column's width, wrapping and
 * number format, rows fitted to their text, header frozen. Steps the API refuses are skipped (a table
 * still works without its look).
 */
export async function formatTable(e: ExcelRef, columns: ColumnLook[]) {
  const { anchor, rows } = await readTable(e);
  const ws = sheetPath(e, anchor.sheet);
  const bottom = anchor.top + Math.max(rows.length, 1);
  const col = (i: number, from = anchor.top, to = bottom) => `${ws}/range(address='${letter(anchor.left + i)}${from}:${letter(anchor.left + i)}${to}')`;
  const t = tablePath(e);
  const steps: BatchRequest[] = [
    { method: "PATCH", url: t, body: { style: "TableStyleLight1", showBandedRows: false, showBandedColumns: false } },
    { method: "PATCH", url: `${t}/headerRowRange/format`, body: { rowHeight: 24, verticalAlignment: "Center", wrapText: false } },
    { method: "PATCH", url: `${t}/headerRowRange/format/fill`, body: { color: "#E5E7EB" } },
    { method: "PATCH", url: `${t}/headerRowRange/format/font`, body: { color: "#111827", bold: true } },
    ...["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"].map((edge) => ({ method: "PATCH" as const, url: `${t}/range/format/borders/${edge}`, body: { style: "Continuous", color: "#9CA3AF", weight: "Thin" } })),
    ...columns.flatMap((c, i): BatchRequest[] => [
      { method: "PATCH", url: `${col(i)}/format`, body: { columnWidth: c.width, wrapText: c.wrap, verticalAlignment: "Top" } },
      ...(c.numberFormat && rows.length ? [{ method: "PATCH" as const, url: col(i, anchor.top + 1, bottom), body: { numberFormat: rows.map(() => [c.numberFormat]) } }] : []),
    ]),
  ];
  // Ten at a time, a refused step skipped.
  for (let i = 0; i < steps.length; i += 10) await Promise.all(steps.slice(i, i + 10).map((s) => graph(s.url, s.method, s.body).catch(() => undefined)));
  // Then the rows fitted to their (wrapped) text, the header frozen.
  await graph(`${t}/dataBodyRange/format/autofitRows`, "POST", {}).catch(() => undefined);
  await graph(`${ws}/freezePanes/freezeRows`, "POST", { count: anchor.top }).catch(() => undefined);
}

/** A table on the first sheet of an empty workbook, with the given header. */
export async function createTable(file: Pick<ExcelRef, "driveId" | "itemId">, header: string[]) {
  const { value } = await graph(`${book(file)}/worksheets?$select=name`);
  const sheet = sheetPath(file, value[0].name);
  const address = `A1:${letter(header.length - 1)}1`;
  await graph(`${sheet}/range(address='${address}')`, "PATCH", { values: [header] });
  return (await graph(`${sheet}/tables/add`, "POST", { address, hasHeaders: true })).name as string;
}

/**
 * URL showing the file in Excel for the web, read only, in an iframe (Graph "preview": the access
 * token is in the URL, valid a few minutes for opening it; no browser session needed).
 */
export async function previewUrl(e: Pick<ExcelRef, "driveId" | "itemId">) {
  const { getUrl } = await graph(`/drives/${e.driveId}/items/${e.itemId}/preview`, "POST", {});
  return getUrl as string;
}
