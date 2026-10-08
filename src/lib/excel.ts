/**
 * Excel tables on SharePoint / OneDrive through the Graph workbook API. Rows are addressed by their
 * position in the table; the fields come from the hidden "_schema" sheet (or are guessed).
 */
import { graph } from "./ms";
import { SCHEMA_HEADER, coerce, fieldToRow, inferFields, isHidden, rowToField, toCell, type Field, type Row } from "./schema";
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

export async function readTable(e: ExcelEntry): Promise<Table> {
  const [range, schema] = await Promise.all([
    graph(`${tablePath(e)}/range?$select=values,valueTypes,numberFormat`),
    graph(`${book(e)}/worksheets('${SCHEMA_SHEET}')/usedRange(valuesOnly=true)?$select=values`).catch(() => null),
  ]);
  const [header, ...data] = range.values as unknown[][];
  const names = header.map(String);
  const declared = new Map<string, Field>(((schema?.values ?? []) as unknown[][]).slice(1).filter((r) => r[0]).map((r) => [String(r[0]), rowToField(r)]));
  const guessed = inferFields(names, (range.valueTypes as string[][]).slice(1), (range.numberFormat as string[][]).slice(1), data);
  const fields = names.map((n, i) => declared.get(n) ?? guessed[i]);
  const rows = data.map((values) =>
    Object.fromEntries(
      fields.map((f, i) => {
        try {
          return [f.name, coerce(f, values[i])];
        } catch {
          return [f.name, values[i]]; // kept as is (shown as invalid)
        }
      }),
    ),
  );
  return { fields, rows, header: names, hasSchema: declared.size > 0 };
}

const check = (t: Table, values: Row) => {
  const out: Row = {};
  for (const [name, value] of Object.entries(values)) {
    const field = t.fields.find((f) => f.name === name);
    if (!field) throw new Error(`Unknown field "${name}" (fields: ${t.fields.map((f) => f.name).join(", ")})`);
    out[name] = coerce(field, value);
  }
  return out;
};
const cells = (t: Table, row: Row) => t.header.map((name) => toCell(row[name]));

/** Partial updates (`values` = changed fields only); returns the new rows. */
export async function updateRows(e: ExcelEntry, t: Table, changes: { index: number; values: Row }[]) {
  for (const { index, values } of changes) {
    if (!t.rows[index]) throw new Error(`No row at index ${index}`);
    const row = { ...t.rows[index], ...check(t, values) };
    await graph(`${tablePath(e)}/rows/itemAt(index=${index})/range`, "PATCH", { values: [cells(t, row)] });
    t.rows[index] = row;
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
  // A hidden id column gets a fresh id.
  const id = t.fields.find(isHidden);
  for (const r of checked) if (id && !r[id.name]) r[id.name] = crypto.randomUUID();
  await graph(`${tablePath(e)}/rows/add`, "POST", { index: null, values: checked.map((r) => cells(t, r)) });
}
const labelOrName = (f: Field) => f.label ?? f.name;

export async function deleteRows(e: ExcelEntry, indexes: number[]) {
  for (const index of [...indexes].sort((a, b) => b - a)) await graph(`${tablePath(e)}/rows/$/itemAt(index=${index})`, "DELETE");
}

/** Writes the fields to the hidden "_schema" sheet and adds the table columns of new fields. */
export async function saveFields(e: ExcelEntry, t: Table, fields: Field[]) {
  const names = new Set<string>();
  for (const f of fields) {
    if (!/^[\p{L}_][\p{L}\p{N}_ ]*$/u.test(f.name)) throw new Error(`Invalid field name "${f.name}"`);
    if (names.has(f.name)) throw new Error(`Duplicate field "${f.name}"`);
    names.add(f.name);
  }
  for (const f of fields.filter((f) => !t.header.includes(f.name))) await graph(`${tablePath(e)}/columns/add`, "POST", { values: [[f.name]] });
  const sheet = `${book(e)}/worksheets('${SCHEMA_SHEET}')`;
  if (!(await graph(`${sheet}?$select=name`).then(() => true, () => false))) {
    await graph(`${book(e)}/worksheets/add`, "POST", { name: SCHEMA_SHEET });
    await graph(sheet, "PATCH", { visibility: "Hidden" });
  }
  await graph(`${sheet}/range(address='A1:G500')/clear`, "POST", { applyTo: "Contents" });
  const values = [[...SCHEMA_HEADER], ...fields.map(fieldToRow)];
  await graph(`${sheet}/range(address='A1:G${values.length}')`, "PATCH", { values });
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

/** Folder that receives the uploaded images: the one chosen in the library, else "<file> images" next to the Excel file. */
export const imagesFolderLabel = (e: ExcelEntry) => (e.imagesFolder ? `the folder "${e.imagesFolder.name}"` : `"${e.name} images", next to the Excel file`);

/** Sharing link of a SharePoint / OneDrive folder → reference. */
export async function resolveFolder(link: string) {
  const item = await graph(`/shares/${shareId(link)}/driveItem?$select=id,name,folder,parentReference`);
  if (!item.folder) throw new Error("This link is not a folder.");
  return { link, driveId: item.parentReference.driveId as string, itemId: item.id as string, name: item.name as string };
}

/** Uploads an image and returns a sharing link for the organization (stored in the cell). */
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
  const link = await graph(`/drives/${target.driveId}/items/${item.id}/createLink`, "POST", { type: "view", scope: "organization" });
  return link.link.webUrl as string;
}
