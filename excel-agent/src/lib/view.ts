import { useEffect, useState } from "react";
import { richToText } from "./rich";
import type { Row, Schema } from "./schema";

export type View = {
  search: string;
  /** Option field → selected values (a row matches any of them). */
  filters: Record<string, string[]>;
  sort: { field: string; direction: "asc" | "desc" } | null;
  layout: "table" | "form";
};

const initial: View = { search: "", filters: {}, sort: null, layout: "table" };

/** View settings of a file, remembered in this browser. */
export function useView(fileId: string) {
  const key = `excel-agent:view:${fileId}`;
  const [view, setView] = useState<View>(() => ({ ...initial, ...JSON.parse(localStorage.getItem(key) ?? "{}") }));
  useEffect(() => localStorage.setItem(key, JSON.stringify(view)), [key, view]);
  return [view, (patch: Partial<View>) => setView((v) => ({ ...v, ...patch }))] as const;
}

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
const plain = (field: { type: string } | undefined, value: unknown) => (field?.type === "text" ? richToText(value) : value);
const flat = (value: unknown) => [value].flat().filter((v) => v !== null && v !== undefined).join(", ");

export function applyView(rows: Row[], view: View, schema: Schema) {
  const search = view.search.trim().toLowerCase();
  let out = rows.filter(
    (row) =>
      Object.entries(view.filters).every(([name, values]) => !values.length || [row[name]].flat().some((v) => values.includes(String(v)))) &&
      (!search || schema.fields.some((f) => flat(plain(f, row[f.name])).toLowerCase().includes(search))),
  );
  if (view.sort) {
    const { field, direction } = view.sort;
    const factor = direction === "asc" ? 1 : -1;
    const def = schema.fields.find((f) => f.name === field);
    out = [...out].sort((a, b) => {
      const [x, y] = [plain(def, a[field]), plain(def, b[field])];
      if (flat(x) === "" || flat(y) === "") return Number(flat(x) === "") - Number(flat(y) === "");
      if (typeof x === "number" && typeof y === "number") return (x - y) * factor;
      return collator.compare(flat(x), flat(y)) * factor;
    });
  }
  return out;
}

export const describeView = (view: View) =>
  [
    view.search && `search "${view.search}"`,
    ...Object.entries(view.filters)
      .filter(([, v]) => v.length)
      .map(([name, values]) => `${name} in (${values.join(", ")})`),
    view.sort && `sorted by ${view.sort.field} ${view.sort.direction}`,
  ]
    .filter(Boolean)
    .join(", ") || "no filter";
