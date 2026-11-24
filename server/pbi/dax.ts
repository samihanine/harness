/**
 * DAX on a semantic model, with the limits that keep models fast for everyone (Settings › Power BI,
 * docs/performance.md): at most `maxRows` rows (TOPN), stopped after `timeoutMs`, a few queries at a time
 * per model, the same query answered from a short cache.
 */
import type { Ref } from "@shared/types";
import { pbi, scope } from "../ms/api";
import { settings } from "../settings";

let PERFORMANCE = { maxRows: 2000, timeoutMs: 20_000, concurrency: 2, intervalMs: 150, cacheMs: 60_000 };
const limits = async () => (PERFORMANCE = { ...PERFORMANCE, ...(await settings()).powerbi });
export type Rows = Record<string, unknown>[];

/** A DAX error with what to do about it (weak models otherwise guess another name, or a wrong cause). */
function daxError(message: string) {
  const hint = /cannot be found|can't be found|not found|introuvable|n'existe pas|cannot find|cannot be determined|doesn't exist/i.test(message)
    ? " → A name does not exist: look up the exact table / column / measure names with search_model or describe_table (never guess or translate them), then fix the query."
    : /^Query \(\d+, \d+\)|syntax/i.test(message)
      ? " → (line, column) of the query: fix the query exactly there."
      : "";
  return new Error(`${message}${hint}`);
}

/* Limits that keep the models fast for everyone (docs/performance.md): PERFORMANCE in src/config.ts. */

/** The rows of a query that were cut at the row limit (more rows exist). */
export type Truncated = Rows & { truncated?: number };

/** The `ORDER BY a DESC, b` items of a query as TOPN arguments ("a, DESC, b, ASC"). */
const orderArgs = (orderBy: string) =>
  splitTop(orderBy)
    .map((item) => {
      const m = item.trim().match(/^([\s\S]*?)\s+(ASC|DESC)$/i);
      return m ? `${m[1]}, ${m[2].toUpperCase()}` : `${item.trim()}, ASC`;
    })
    .join(", ");

/** Splits at the commas outside parentheses, brackets and quotes. */
function splitTop(text: string) {
  const parts: string[] = [];
  let depth = 0;
  let quote = "";
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quote) quote = c === quote ? "" : quote;
    else if (c === '"' || c === "'") quote = c;
    else if (c === "(" || c === "[") depth++;
    else if (c === ")" || c === "]") depth--;
    else if (c === "," && depth === 0) (parts.push(text.slice(from, i)), (from = i + 1));
  }
  return [...parts, text.slice(from)];
}

/**
 * The query with its last EVALUATE limited to `limit` rows (TOPN, the ORDER BY kept as its order):
 * the model computes no more rows than the app shows. Undefined when the query has another shape.
 */
export function limitRows(query: string, limit: number) {
  const at = [...query.matchAll(/\bEVALUATE\b/gi)].at(-1)?.index;
  if (at === undefined) return undefined;
  const rest = query.slice(at + "EVALUATE".length);
  const order = rest.search(/\bORDER\s+BY\b/i);
  if (/\bSTART\s+AT\b/i.test(rest)) return undefined;
  const body = (order < 0 ? rest : rest.slice(0, order)).trim().replace(/;$/, "");
  if (!body || /^ROW\s*\(/i.test(body) || /^\{/.test(body)) return undefined;
  const orderBy = order < 0 ? "" : rest.slice(order).replace(/^ORDER\s+BY\b/i, "").trim().replace(/;$/, "");
  return `${query.slice(0, at)}EVALUATE\nTOPN(${limit}, ${body}${orderBy ? `, ${orderArgs(orderBy)}` : ""})${orderBy ? `\nORDER BY ${orderBy}` : ""}`;
}

/** At most `concurrency` queries at a time per model, `intervalMs` apart (the model's capacity is shared). */
const queues = new Map<string, { running: number; last: number; waiting: (() => void)[] }>();
async function slot(model: string) {
  const q = queues.get(model) ?? queues.set(model, { running: 0, last: 0, waiting: [] }).get(model)!;
  while (q.running >= Math.max(1, PERFORMANCE.concurrency)) await new Promise<void>((r) => q.waiting.push(r));
  q.running++;
  const wait = q.last + PERFORMANCE.intervalMs - Date.now();
  if (wait > 0) await new Promise((r) => setTimeout(r, wait));
  q.last = Date.now();
  return () => {
    q.running--;
    q.waiting.shift()?.();
  };
}

/** The same query on the same model within `cacheMs`: the rows already read. */
const cache = new Map<string, { at: number; rows: Promise<Truncated> }>();

/**
 * Runs a DAX query: at most `maxRows` rows (more are cut, `truncated` says how many were asked for at
 * least), stopped after `timeoutMs`, queued per model and cached a little (src/config.ts PERFORMANCE).
 */
export async function dax(dataset: Ref, query: string): Promise<Truncated> {
  const key = `${dataset.id}\n${query}`;
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < PERFORMANCE.cacheMs) return hit.rows;
  const rows = runDax(dataset, query);
  cache.set(key, { at: Date.now(), rows });
  rows.catch(() => cache.delete(key));
  for (const [k, v] of cache) if (Date.now() - v.at > PERFORMANCE.cacheMs) cache.delete(k);
  return rows;
}

async function runDax(dataset: Ref, query: string): Promise<Truncated> {
  await limits();
  const max = Math.max(1, PERFORMANCE.maxRows);
  const limited = limitRows(query, max + 1);
  const done = await slot(dataset.id);
  try {
    let rows: Rows;
    try {
      rows = await executeDax(dataset, limited ?? query);
    } catch (e) {
      // The limit made the query invalid (an unusual shape): asked again as written, cut here.
      if (!limited || /took longer/.test(String(e))) throw e;
      rows = await executeDax(dataset, query);
    }
    const out: Truncated = rows.slice(0, max);
    if (rows.length > max) out.truncated = rows.length;
    return out;
  } finally {
    done();
  }
}

async function executeDax(dataset: Ref, query: string): Promise<Rows> {
  const abort = new AbortController();
  const timer = setTimeout(() => abort.abort(), PERFORMANCE.timeoutMs);
  const data = await pbi(`${scope(dataset.groupId)}/datasets/${dataset.id}/executeQueries`, "POST", { queries: [{ query }], serializerSettings: { includeNulls: true } }, abort.signal)
    .catch((e) => {
      if (abort.signal.aborted) throw new Error(`The query took longer than ${PERFORMANCE.timeoutMs / 1000} s and was stopped (limit of the app, to keep the model fast). Make it lighter: fewer columns, a filter, SUMMARIZECOLUMNS rather than FILTER on a whole table, or count the rows first.`);
      throw daxError(e instanceof Error ? e.message : String(e));
    })
    .finally(() => clearTimeout(timer));
  const result = data.results[0];
  if (result.error) throw daxError(result.error.message);
  return (result.tables?.[0]?.rows ?? []).map((row: Record<string, unknown>) =>
    Object.fromEntries(Object.entries(row).map(([k, v]) => [k.replace(/^.*\[(.*)\]$/, "$1"), v])),
  );
}

/** Refreshes on a Pro (shared) capacity: 8 a day per model, by API and scheduled together (refreshes started in Power BI itself are not counted). */
export const REFRESH_QUOTA = 8;

/** The refreshes that count in the daily quota (by API or scheduled) over the last 24 hours. */
export async function refreshesToday(dataset: Ref) {
  const { value } = await pbi<{ value: { refreshType: string; startTime: string }[] }>(`${scope(dataset.groupId)}/datasets/${dataset.id}/refreshes?$top=30`);
  const since = Date.now() - 24 * 3600 * 1000;
  return value.filter((r) => (r.refreshType === "ViaApi" || r.refreshType === "Scheduled") && Date.parse(r.startTime) >= since).length;
}

/* ----------------------------- Column values ----------------------------- */

const columnRef = (table: string, column: string) => `'${table.replace(/'/g, "''")}'[${column.replace(/]/g, "]]")}]`;
const literal = (v: unknown) => (typeof v === "number" ? String(v) : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : `"${String(v).replace(/"/g, '""')}"`);

/** Distinct values of a column (first `VALUES_SHOWN`), kept for the session: what slicers and filters can take. */
const VALUES_SHOWN = 40;
const valuesCache = new Map<string, Promise<{ values: unknown[]; more: boolean }>>();
export function columnValues(dataset: Ref, table: string, column: string) {
  const key = `${dataset.id}|${table}|${column}`;
  if (!valuesCache.has(key)) {
    const run = dax(dataset, `EVALUATE TOPN(${VALUES_SHOWN + 1}, VALUES(${columnRef(table, column)}))`).then((rows) => {
      const values = rows.map((r) => Object.values(r)[0]).filter((v) => v !== null && v !== "");
      return { values: values.slice(0, VALUES_SHOWN), more: values.length > VALUES_SHOWN };
    });
    valuesCache.set(key, run);
    run.catch(() => valuesCache.delete(key));
  }
  return valuesCache.get(key)!;
}

/**
 * The values to filter on, as the model holds them: matched without case ("fall" → "Fall"), and
 * refused when they do not exist (the error lists the valid ones), so the AI never filters on a value
 * that empties the visuals. Without access to the model (no Build right), kept as given.
 */
export async function checkValues(dataset: Ref | undefined, table: string, column: string, given: unknown[]) {
  if (!dataset || !given.length) return given;
  const found = await dax(dataset, `EVALUATE FILTER(VALUES(${columnRef(table, column)}), ${columnRef(table, column)} IN {${given.map(literal).join(", ")}})`)
    .then((rows) => rows.map((r) => Object.values(r)[0]))
    .catch(() => undefined);
  if (!found) return given;
  const same = (a: unknown, b: unknown) => String(a).toLowerCase() === String(b).toLowerCase();
  const missing = given.filter((g) => !found.some((f) => same(f, g)));
  if (missing.length) {
    const { values, more } = await columnValues(dataset, table, column).catch(() => ({ values: [] as unknown[], more: false }));
    throw new Error(
      `${table}[${column}] has no value ${missing.map((m) => JSON.stringify(m)).join(", ")}. ${values.length ? `Its values${more ? " (first " + values.length + ")" : ""}: ${values.map((v) => JSON.stringify(v)).join(", ")}.` : ""} Pick the matching one (e.g. a translation or synonym of the user's word) and retry${more ? ", or search them with run_dax" : ""}.`,
    );
  }
  return given.map((g) => found.find((f) => same(f, g)) ?? g);
}

/** Rows as a compact text table for the AI. */
export function rowsText(rows: Truncated, max = 50) {
  if (rows.length === 0) return "(no rows)";
  const cut = rows.truncated ? `\n⚠ Cut at the app's limit of ${rows.length} rows (the query returns more): aggregate, filter or TOPN; to know how many, EVALUATE ROW("n", COUNTROWS(<the same table>)).` : "";
  const columns = Object.keys(rows[0]);
  // Long values (e.g. base64 images) shown by their start and length: one would fill the AI's context.
  const cell = (v: unknown) => {
    const text = String(v ?? "");
    // 32,766 = Power BI's limit for a text value: a value of exactly that length was cut.
    return text.length > 120 ? `${text.slice(0, 60)}… (${text.length} characters${text.length === 32766 ? ", CUT at Power BI's limit" : ""})` : text;
  };
  const lines = rows.slice(0, max).map((r) => columns.map((c) => cell(r[c])).join(" | "));
  return [columns.join(" | "), ...lines, rows.length > max ? `… ${rows.length - max} more rows` : ""].filter(Boolean).join("\n") + cut;
}

