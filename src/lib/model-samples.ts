/**
 * Examples of a model's data for the AI (with the right to query it): a few values of each column,
 * from random rows, and each measure's total. Short, cached a few minutes: it saves the model calls
 * that would only look at the data to learn its shape (formats, codes, dates…).
 */
import { dax, type Ref } from "./pbi";
import type { ModelInfo } from "./model";

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
