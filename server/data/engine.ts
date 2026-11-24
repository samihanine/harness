/**
 * The Data Scientist's engine: an in-memory SQL database (SQLite, in Bun) per conversation, where DAX
 * results and Excel / CSV sheets become tables the AI joins, filters and aggregates freely with SQL (CTEs,
 * window functions, CASE…). Nothing is stored: the database is rebuilt from the analysis document when
 * the server no longer has it.
 */
import { Database } from "bun:sqlite";
import readXlsxFile from "read-excel-file/node";
import type { Analysis, Rows, Source } from "@shared/analysis";
import { dax } from "../pbi/dax";
import { readStoredFile } from "../files";

export const MAX_ROWS = 200_000;

export type Engine = { db: Database; loaded: Map<string, { source: Source; columns: { name: string; type: string }[]; rows: number }>; signature: string };

export const ident = (name: string) => `"${name.replace(/"/g, '""')}"`;
export const safeName = (name: string) => name.trim().replace(/[^\w]+/g, "_").replace(/^(\d)/, "_$1").slice(0, 60) || "t";

/** The column types of rows: REAL when every filled value is a number, else TEXT. */
function columnsOf(rows: Rows) {
  const names = [...new Set(rows.flatMap((r) => Object.keys(r)))];
  return names.map((name) => {
    const values = rows.map((r) => r[name]).filter((v) => v !== null && v !== undefined && v !== "");
    const numeric = values.length > 0 && values.every((v) => typeof v === "number" || typeof v === "boolean");
    return { name, type: numeric ? "REAL" : "TEXT" };
  });
}

const cellValue = (v: unknown) => (v instanceof Date ? v.toISOString().replace("T00:00:00.000Z", "") : typeof v === "boolean" ? (v ? 1 : 0) : v === undefined ? null : typeof v === "object" && v !== null ? JSON.stringify(v) : v);

export function createTable(engine: Engine, name: string, rows: Rows, source: Source) {
  const cols = columnsOf(rows);
  if (!cols.length) throw new Error(`"${name}" has no column (no rows returned).`);
  const t = ident(name);
  engine.db.run(`DROP VIEW IF EXISTS ${t}`);
  engine.db.run(`DROP TABLE IF EXISTS ${t}`);
  engine.db.run(`CREATE TABLE ${t} (${cols.map((c) => `${ident(c.name)} ${c.type}`).join(", ")})`);
  const insert = engine.db.prepare(`INSERT INTO ${t} VALUES (${cols.map(() => "?").join(", ")})`);
  engine.db.transaction(() => {
    for (const r of rows) insert.run(...(cols.map((c) => cellValue(r[c.name])) as never[]));
  })();
  engine.loaded.set(name, { source, columns: cols, rows: rows.length });
  return cols;
}

/** The rows of a DAX source (column names without their table prefix). */
export async function daxRows(source: Extract<Source, { kind: "dax" }>) {
  const rows = await dax({ id: source.datasetId, groupId: source.groupId }, source.query);
  return rows as Rows;
}

/** Sheets of an uploaded file: Excel (every sheet) or CSV / TSV (one). The first row holds the column names. */
export async function fileSheets(fileId: string): Promise<{ name: string; header: string[]; rows: Rows }[]> {
  const { ref, data } = await readStoredFile(fileId);
  if (/\.(csv|tsv|txt)$/i.test(ref.name) || /csv|text\/plain/.test(ref.type)) {
    const text = new TextDecoder().decode(data).replace(/^﻿/, "");
    const first = text.split(/\r?\n/, 1)[0];
    const sep = [";", "\t", ",", "|"].sort((a, b) => first.split(b).length - first.split(a).length)[0];
    const lines = parseCsv(text, sep);
    return [toSheet(ref.name.replace(/\.\w+$/, ""), lines)];
  }
  const sheets = await readXlsxFile(Buffer.from(data));
  return sheets.map((s) => toSheet(s.sheet, s.data as unknown[][]));
}

function toSheet(name: string, lines: unknown[][]) {
  // The header: the first row with at least two filled cells (titles above a table are skipped).
  const start = Math.max(0, lines.findIndex((l) => l.filter((c) => c !== null && c !== "").length >= 2));
  const header = (lines[start] ?? []).map((h, i) => String(h ?? "").trim() || `Column${i + 1}`);
  const seen = new Map<string, number>();
  const unique = header.map((h) => {
    const n = (seen.get(h) ?? 0) + 1;
    seen.set(h, n);
    return n > 1 ? `${h}_${n}` : h;
  });
  const rows = lines
    .slice(start + 1)
    .filter((l) => l.some((c) => c !== null && c !== ""))
    .slice(0, MAX_ROWS)
    .map((l) => Object.fromEntries(unique.map((h, i) => [h, numberish(l[i])])));
  return { name, header: unique, rows };
}

/** "1 234,5" / "12%" in a CSV → numbers; other values as they are. */
function numberish(v: unknown) {
  if (typeof v !== "string") return v ?? null;
  const t = v.trim();
  if (/^-?[\d\s ]+([.,]\d+)?%?$/.test(t) && /\d/.test(t)) {
    const n = Number(t.replace(/[\s ]/g, "").replace(",", ".").replace("%", ""));
    if (!Number.isNaN(n)) return t.endsWith("%") ? n / 100 : n;
  }
  return t === "" ? null : t;
}

function parseCsv(text: string, sep: string): string[][] {
  const out: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') (cell += '"'), i++;
      else if (c === '"') quoted = false;
      else cell += c;
    } else if (c === '"') quoted = true;
    else if (c === sep) row.push(cell), (cell = "");
    else if (c === "\n" || c === "\r") {
      if (c === "\r" && text[i + 1] === "\n") i++;
      row.push(cell);
      out.push(row);
      row = [];
      cell = "";
    } else cell += c;
  }
  if (cell || row.length) row.push(cell), out.push(row);
  return out;
}

export async function loadSource(engine: Engine, source: Source) {
  if (source.kind === "dax") return createTable(engine, source.name, await daxRows(source), source);
  const sheets = await fileSheets(source.fileId);
  const sheet = source.sheet ? sheets.find((s) => s.name.toLowerCase() === source.sheet!.toLowerCase()) : sheets[0];
  if (!sheet) throw new Error(`No sheet "${source.sheet}" in ${source.fileName}. Sheets: ${sheets.map((s) => s.name).join(", ")}`);
  return createTable(engine, source.name, sheet.rows, source);
}

const engines = new WeakMap<object, Engine>();

/** The conversation's database, built (again) from the analysis when missing or changed elsewhere. */
export async function engineFor(scratch: Record<string, any>, analysis: Analysis | undefined) {
  let engine: Engine | undefined = scratch.engine;
  if (!engine) {
    engine = { db: new Database(":memory:"), loaded: new Map(), signature: "" };
    scratch.engine = engine;
  }
  if (analysis) {
    for (const s of analysis.sources) {
      const loaded = engine.loaded.get(s.name);
      if (!loaded || JSON.stringify(loaded.source) !== JSON.stringify({ ...s, rows: loaded.source.rows })) await loadSource(engine, s).catch(() => undefined);
    }
    for (const step of analysis.steps) {
      const exists = engine.db.query(`SELECT name FROM sqlite_master WHERE name = ?`).get(step.name);
      if (!exists) {
        try {
          engine.db.run(`CREATE VIEW ${ident(step.name)} AS ${step.sql}`);
        } catch {}
      }
    }
  }
  void engines;
  return engine;
}

/** Runs a read query (SELECT / WITH). */
export function select(engine: Engine, sql: string, limit = 1000): { columns: string[]; rows: Rows; truncated: boolean } {
  const q = sql.trim().replace(/;+\s*$/, "");
  if (!/^(select|with)\b/i.test(q)) throw new Error("Only SELECT (or WITH … SELECT) queries here: create steps with create_step.");
  if (/;\s*\S/.test(q)) throw new Error("One query at a time.");
  const stmt = engine.db.query(q);
  const rows = stmt.all() as Rows;
  return { columns: stmt.columnNames, rows: rows.slice(0, limit), truncated: rows.length > limit };
}

/** The tables and views of the database with their columns. */
export function schemaText(engine: Engine) {
  const objects = engine.db.query(`SELECT name, type FROM sqlite_master WHERE type IN ('table','view') ORDER BY type, name`).all() as { name: string; type: string }[];
  if (!objects.length) return "No table loaded yet: load sources with load_dax / load_file.";
  return objects
    .map((o) => {
      const cols = engine.db.query(`SELECT name, type FROM pragma_table_info(?)`).all(o.name) as { name: string; type: string }[];
      const count = (engine.db.query(`SELECT COUNT(*) AS n FROM ${ident(o.name)}`).get() as { n: number } | null)?.n ?? 0;
      return `- ${o.type} ${o.name} (${count} rows): ${cols.map((c) => `${c.name}${c.type ? ` ${c.type.toLowerCase()}` : ""}`).join(", ")}`;
    })
    .join("\n");
}

/** Rows as a short text table. */
export function rowsText(columns: string[], rows: Rows, max = 60) {
  if (!rows.length) return "(no rows)";
  const cell = (v: unknown) => {
    const t = v === null || v === undefined ? "" : typeof v === "number" ? String(Math.round(v * 10000) / 10000) : String(v);
    return t.length > 80 ? `${t.slice(0, 77)}…` : t;
  };
  return [columns.join(" | "), ...rows.slice(0, max).map((r) => columns.map((c) => cell(r[c])).join(" | ")), rows.length > max ? `… ${rows.length - max} more rows` : ""].filter(Boolean).join("\n");
}
