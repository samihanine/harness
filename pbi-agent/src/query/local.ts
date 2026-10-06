/**
 * Local semantic models built from Excel files. Each sheet (or Excel table) becomes a model
 * table with typed columns; relationships are inferred on matching key columns; measures are
 * SQL aggregates. Queries run in the browser with AlaSQL, in SQL (the local query language).
 */
import alasql from "alasql";
import ExcelJS from "exceljs";
import { createStore, get, set } from "idb-keyval";
import type { QueryResult } from "@/pbi/api";
import type { ModelContent, TableInfo } from "@/pbi/types";

const files = createStore("pbi-agent-files", "handles");
export const saveDatasetFile = (id: string, handle: FileSystemFileHandle) => set(id, handle, files);
const datasetFile = (id: string) => get<FileSystemFileHandle>(id, files);

type Cell = string | number | boolean | Date | null;
type Sheet = { name: string; columns: string[]; rows: Record<string, Cell>[] };

const plain = (value: ExcelJS.CellValue): Cell => {
  if (value === null || value === undefined) return null;
  if (value instanceof Date || typeof value !== "object") return value as Cell;
  if ("richText" in value) return value.richText.map((r) => r.text).join("");
  if ("result" in value) return plain(value.result as ExcelJS.CellValue);
  if ("text" in value) return String(value.text);
  return null;
};

/** Data of every sheet: an Excel table when there is one, else the block starting at A1. */
function readSheets(workbook: ExcelJS.Workbook): Sheet[] {
  const sheets: Sheet[] = [];
  workbook.eachSheet((ws) => {
    // Loaded tables expose tableRef, tables added in this session ref.
    const tables = Object.values((ws as unknown as { tables: Record<string, { table: { ref?: string; tableRef?: string; name: string } }> }).tables ?? {});
    const ranges = tables.length
      ? tables.map((t) => ({ name: t.table.name, ref: (t.table.tableRef ?? t.table.ref)! }))
      : ws.actualRowCount > 1
        ? [{ name: ws.name, ref: `A1:${ws.getColumn(ws.actualColumnCount).letter}${ws.rowCount}` }]
        : [];
    for (const range of ranges) {
      const [start, end] = range.ref.split(":").map((a) => ws.getCell(a));
      const columns: string[] = [];
      for (let c = Number(start.col); c <= Number(end.col); c++) columns.push(String(plain(ws.getCell(Number(start.row), c).value) ?? `Column${c}`).trim());
      const rows: Record<string, Cell>[] = [];
      for (let r = Number(start.row) + 1; r <= Number(end.row); r++) {
        const row: Record<string, Cell> = {};
        let empty = true;
        columns.forEach((name, i) => {
          const v = plain(ws.getCell(r, Number(start.col) + i).value);
          if (v !== null && v !== "") empty = false;
          row[name] = v;
        });
        if (!empty) rows.push(row);
      }
      sheets.push({ name: range.name, columns, rows });
    }
  });
  return sheets;
}

function typeOf(values: Cell[]) {
  const present = values.filter((v) => v !== null && v !== "");
  if (present.length === 0) return "String";
  if (present.every((v) => v instanceof Date)) return "DateTime";
  if (present.every((v) => typeof v === "boolean")) return "Boolean";
  if (present.every((v) => typeof v === "number")) return present.every((v) => Number.isInteger(v)) ? "Int64" : "Double";
  return "String";
}

/**
 * Builds the model of an Excel file. Sheets that are not Excel tables are turned into tables
 * in the file (so it can also be used as a Power BI source); `convert` = false to leave it as is.
 */
export async function importExcel(handle: FileSystemFileHandle, { convert = true } = {}) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await (await handle.getFile()).arrayBuffer());
  const sheets = readSheets(workbook);
  if (sheets.length === 0) throw new Error("No data found in this file (a header row is expected).");

  let converted = 0;
  if (convert)
    workbook.eachSheet((ws) => {
      const tables = Object.keys((ws as unknown as { tables?: Record<string, unknown> }).tables ?? {});
      const sheet = sheets.find((s) => s.name === ws.name);
      if (tables.length || !sheet) return;
      ws.spliceRows(1, ws.rowCount);
      ws.addTable({
        name: ws.name.replace(/[^\w]/g, "_").replace(/^(\d)/, "_$1"),
        ref: "A1",
        headerRow: true,
        style: { theme: "TableStyleLight1", showRowStripes: true },
        columns: sheet.columns.map((name) => ({ name, filterButton: true })),
        rows: sheet.rows.map((row) => sheet.columns.map((c) => row[c] as ExcelJS.CellValue)),
      });
      converted++;
    });
  if (converted) {
    const stream = await handle.createWritable();
    await stream.write(await workbook.xlsx.writeBuffer());
    await stream.close();
  }

  const tables: TableInfo[] = sheets.map((sheet) => ({
    name: sheet.name,
    columns: sheet.columns.map((name) => ({ name, dataType: typeOf(sheet.rows.map((r) => r[name])) })),
    measures: [],
  }));
  return { model: { tables, relationships: inferRelationships(sheets), warnings: [] as string[], methods: ["excel"], extractedAt: new Date().toISOString() } satisfies ModelContent, converted, sheets: sheets.map((s) => s.name) };
}

/** Many-to-one when a column of one table matches a unique column of the same name in another. */
function inferRelationships(sheets: Sheet[]) {
  const out: ModelContent["relationships"] = [];
  for (const one of sheets)
    for (const key of one.columns) {
      const values = one.rows.map((r) => r[key]);
      if (values.length === 0 || new Set(values).size !== values.length) continue;
      for (const many of sheets)
        if (many !== one && many.columns.includes(key) && many.rows.length > one.rows.length)
          out.push({ from: `'${many.name}'[${key}]`, to: `'${one.name}'[${key}]`, active: true });
    }
  return out;
}

/* --------------------------------- queries -------------------------------- */

const loaded = new Map<string, { modified: number; db: InstanceType<typeof alasql.Database> }>();

/** AlaSQL database of a local dataset, reloaded when the file changed. */
async function database(datasetId: string) {
  const handle = await datasetFile(datasetId);
  if (!handle) throw new Error("The Excel file of this dataset is not available in this browser: add it again.");
  if ((await handle.queryPermission({ mode: "read" })) !== "granted" && (await handle.requestPermission({ mode: "read" })) !== "granted")
    throw new Error("Access to the Excel file was refused.");
  const file = await handle.getFile();
  const cached = loaded.get(datasetId);
  if (cached && cached.modified === file.lastModified) return cached.db;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const db = new alasql.Database();
  for (const sheet of readSheets(workbook)) {
    db.exec(`CREATE TABLE [${sheet.name}]`);
    (db.tables as Record<string, { data: unknown[] }>)[sheet.name].data = sheet.rows;
  }
  loaded.set(datasetId, { modified: file.lastModified, db });
  return db;
}

/** SQL on a local dataset: tables and columns in [brackets], e.g. SELECT [Region], SUM([Amount]) FROM [Sales] GROUP BY [Region]. */
export async function executeSql(datasetId: string, sql: string): Promise<QueryResult> {
  const db = await database(datasetId);
  const result = db.exec(sql) as unknown;
  const rows = (Array.isArray(result) ? (Array.isArray(result.at(-1)) ? result.at(-1) : result) : []) as Record<string, unknown>[];
  return { columns: rows[0] ? Object.keys(rows[0]) : [], rows };
}
