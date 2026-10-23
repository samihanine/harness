/**
 * Examples of a model's data for the AI (with the right to query it): a few values of each column,
 * from random rows, and each measure's total. Short, cached a few minutes: it saves the model calls
 * that would only look at the data to learn its shape (formats, codes, dates…).
 */
import { dax, type Ref } from "./pbi";
import { relevantTables, type ModelInfo } from "./model";

const TTL = 5 * 60_000;
const MAX = 4_000;
const cache = new Map<string, { at: number; text: Promise<string> }>();
const q = (name: string) => `'${name.replace(/'/g, "''")}'`;
const b = (name: string) => `[${name.replace(/]/g, "]]")}]`;

/** A value as the AI reads it: dates without midnight, numbers rounded, long texts cut. */
const show = (v: unknown) => {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(Math.round(v * 100) / 100);
  return String(v).replace(/T00:00:00(\.000)?Z?$/, "").slice(0, 32);
};
const timeout = <T,>(p: Promise<T>, ms: number) => Promise.race([p, new Promise<never>((_, j) => setTimeout(() => j(new Error("timeout")), ms))]);

async function read(ref: Ref, info: ModelInfo) {
  const tables = info.tables.filter((t) => !t.hidden).slice(0, 16);
  const lines = await Promise.all(
    tables.map(async (t) => {
      const cols = t.columns.filter((c) => !c.hidden).slice(0, 14);
      if (!cols.length) return "";
      // Random rows: TOPN ordered by RAND() picks different rows at each read.
      const query = `EVALUATE TOPN(5, SELECTCOLUMNS(${q(t.name)}, ${cols.map((c, i) => `"c${i}", ${q(t.name)}${b(c.name)}`).join(", ")}), RAND())`;
      const rows = await timeout(dax(ref, query), 10_000).catch(() => undefined);
      if (!rows?.length) return "";
      const values = cols.map((c, i) => {
        const distinct = [...new Set(rows.map((r) => show(r[`[c${i}]`] ?? r[`c${i}`])).filter(Boolean))].slice(0, 3);
        return distinct.length ? `${c.name}: ${distinct.join(" | ")}` : "";
      });
      return `- '${t.name}': ${values.filter(Boolean).join("; ")}`;
    }),
  );
  const measures = info.tables.flatMap((t) => t.measures.filter((m) => !m.hidden)).slice(0, 40);
  let totals = "";
  if (measures.length) {
    const row = await timeout(dax(ref, `EVALUATE ROW(${measures.map((m, i) => `"m${i}", ${b(m.name)}`).join(", ")})`), 10_000).catch(() => undefined);
    if (row?.[0]) totals = `Measures over all the data: ${measures.map((m, i) => `[${m.name}] = ${show(row[0][`[m${i}]`] ?? row[0][`m${i}`]) || "blank"}`).join("; ")}`;
  }
  const text = ["Examples of the data (random rows; exact values as stored):", ...lines.filter(Boolean), totals].filter(Boolean).join("\n");
  return text.length > MAX ? `${text.slice(0, MAX)}\n…` : text;
}

/** The examples of a model (empty when it cannot be queried). */
export function modelSamples(ref: Ref, info: ModelInfo | undefined): Promise<string> {
  if (!info) return Promise.resolve("");
  const key = `${ref.id}|${info.readAt}`;
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < TTL) return hit.text;
  const text = read(ref, info).catch(() => "");
  cache.set(key, { at: Date.now(), text });
  return text;
}

/* ------------------------------- Value profile ------------------------------ */

export type ProfileProgress = (label: string, done: number, total: number) => void;
const TEXT = /string|text/i;

/** Runs tasks a few at a time (the API throttles bursts). */
async function pool<T>(tasks: (() => Promise<T>)[], size: number) {
  const out: T[] = [];
  let next = 0;
  await Promise.all(Array.from({ length: Math.min(size, tasks.length) }, async () => {
    while (next < tasks.length) {
      const i = next++;
      out[i] = await tasks[i]();
    }
  }));
  return out;
}

/**
 * The value profile of a model (Datasets › "Value examples", on demand): for each visible column its
 * number of distinct values, up to 100 random values (text) or its range (numbers, dates), and each
 * measure's total. Few queries: one per table for counts and ranges, one per text column for values,
 * measures by twenty.
 */
export async function profileModel(ref: Ref, info: ModelInfo, progress?: ProfileProgress): Promise<ModelInfo> {
  const next: ModelInfo = structuredClone(info);
  const tables = next.tables.filter((t) => !t.hidden);
  const texts = tables.flatMap((t) => t.columns.filter((c) => !c.hidden && (!c.type || TEXT.test(c.type))).map((c) => ({ t, c })));
  const measures = tables.flatMap((t) => t.measures.filter((m) => !m.hidden));
  const chunks = Array.from({ length: Math.ceil(measures.length / 20) }, (_, i) => measures.slice(i * 20, i * 20 + 20));
  const total = tables.length + texts.length + chunks.length;
  let done = 0;
  const step = (label: string) => progress?.(label, ++done, total);
  progress?.("Counting the values", 0, total);

  // 1. Distinct counts and ranges: one query per table.
  await pool(
    tables.map((t) => async () => {
      const cols = t.columns.filter((c) => !c.hidden);
      if (cols.length) {
        const parts = cols.flatMap((c, i) => [`"d${i}", DISTINCTCOUNT(${q(t.name)}${b(c.name)})`, ...(TEXT.test(c.type ?? "string") ? [] : [`"a${i}", MIN(${q(t.name)}${b(c.name)})`, `"z${i}", MAX(${q(t.name)}${b(c.name)})`])]);
        const row = (await timeout(dax(ref, `EVALUATE ROW(${parts.join(", ")})`), 30_000).catch(() => undefined))?.[0];
        if (row)
          cols.forEach((c, i) => {
            const n = Number(row[`[d${i}]`] ?? row[`d${i}`]);
            const lo = show(row[`[a${i}]`] ?? row[`a${i}`]);
            const hi = show(row[`[z${i}]`] ?? row[`z${i}`]);
            if (!TEXT.test(c.type ?? "string")) c.values = `${n} distinct, from ${lo || "blank"} to ${hi || "blank"}`;
            else if (Number.isFinite(n)) c.values = `${n} distinct`;
          });
      }
      step(`Counting the values of '${t.name}'`);
    }),
    4,
  );
  // 2. Random values of text columns (all of them when 100 or fewer).
  await pool(
    texts.map(({ t, c }) => async () => {
      const rows = await timeout(dax(ref, `EVALUATE TOPN(100, VALUES(${q(t.name)}${b(c.name)}), RAND())`), 30_000).catch(() => undefined);
      if (rows?.length) {
        const values = [...new Set(rows.map((r) => show(Object.values(r)[0])).filter(Boolean))];
        const n = Number(String(c.values ?? "").match(/^(\d+) distinct/)?.[1] ?? values.length);
        c.values = `${n} distinct${n > values.length ? `, ${values.length} random` : ""}: ${values.join(" | ")}`;
      }
      step(`Values of '${t.name}'[${c.name}]`);
    }),
    4,
  );
  // 3. Measures over all the data, by twenty (a measure failing alone does not lose the others).
  await pool(
    chunks.map((list) => async () => {
      const row = (await timeout(dax(ref, `EVALUATE ROW(${list.map((m, i) => `"m${i}", ${b(m.name)}`).join(", ")})`), 30_000).catch(() => undefined))?.[0];
      if (row) list.forEach((m, i) => (m.total = show(row[`[m${i}]`] ?? row[`m${i}`]) || "blank"));
      step("Measures");
    }),
    2,
  );
  next.profiledAt = new Date().toISOString();
  return next;
}

/** The profile in short for the AI's context (the full values are in describe_table). */
export function profileText(info: ModelInfo, focus?: string) {
  // The tables of the request first: a long profile is cut at the end.
  const relevant = relevantTables(info, focus);
  const lines = info.tables
    .filter((t) => !t.hidden)
    .sort((a, b) => Number(relevant.has(b.name)) - Number(relevant.has(a.name)))
    .map((t) => {
      const cols = t.columns.filter((c) => !c.hidden && c.values).map((c) => `${c.name}: ${c.values!.length > 160 ? `${c.values!.slice(0, 160)}…` : c.values}`);
      const ms = t.measures.filter((m) => !m.hidden && m.total).map((m) => `[${m.name}] = ${m.total}`);
      return cols.length || ms.length ? `- '${t.name}': ${[...cols, ...ms].join("; ")}` : "";
    })
    .filter(Boolean);
  const text = lines.join("\n");
  // A big model keeps a short context: the rest stays in describe_table.
  return lines.length ? `Value profile (exact values as stored; all of them with describe_table):\n${text.length > 6000 ? `${text.slice(0, 6000)}\n…` : text}` : "";
}
