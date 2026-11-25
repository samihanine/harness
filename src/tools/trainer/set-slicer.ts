/**
 * Changes a slicer of the user's report: values (checked against the model), a range, paths of a hierarchy
 * slicer, or nothing selected. A hierarchy slicer (Year › Month › Day, Zone › SubZone) takes no plain or
 * range filter: its selection is a tree, built here — a date range becomes the dates' years, months and days
 * read from the model. When a slicer refuses any selection, the same filter goes on the page only after the
 * slicer is cleared (a page filter against the slicer's old selection empties the visuals).
 */
import { z } from "zod";
import { checkValues, dax } from "@server/pbi/dax";
import { modelInfo } from "@server/pbi/models";
import type { ModelInfo } from "@server/pbi/model";
import type { Ref } from "@shared/types";
import { tool, type ToolContext } from "../tool";
import { viewText, type ViewState } from "./view-text";

type Target = { table: string; column?: string; measure?: string; hierarchy?: string; hierarchyLevel?: string };
type Slicer = ViewState["slicers"][number] & { targets?: Target[] };
type Node = { value: unknown; operator: "Selected" | "Inherited"; children?: Node[] };

const col = (t: string, c: string) => `'${t.replace(/'/g, "''")}'[${c.replace(/]/g, "]]")}]`;
const DATE_PARTS: Record<string, (d: string) => string> = {
  year: (d) => `YEAR(${d})`,
  quarter: (d) => `"Qtr " & QUARTER(${d})`,
  // The automatic date hierarchy names months in English, whatever the model's language.
  month: (d) => `SWITCH(MONTH(${d}), ${["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"].map((m, i) => `${i + 1}, "${m}"`).join(", ")})`,
  day: (d) => `DAY(${d})`,
};

/** The column of each level of a hierarchy slicer (or how to compute it from a date, for automatic date hierarchies). */
function levelsOf(targets: Target[], info: ModelInfo | undefined) {
  const table = info?.tables.find((t) => t.name === targets[0]?.table);
  const hierarchy = targets[0]?.hierarchy && table?.hierarchies.map((h) => h.match(/^(.+) \((.+)\)$/)).find((m) => m?.[1] === targets[0].hierarchy);
  const columns = hierarchy ? hierarchy[2].split(" › ") : [];
  return targets.map((t, i) => {
    const name = t.column ?? t.hierarchyLevel ?? "";
    const real = table?.columns.find((c) => c.name === name)?.name ?? (columns.length === targets.length ? columns[i] : undefined);
    return { table: t.table, column: real, part: real ? undefined : DATE_PARTS[name.toLowerCase()], name };
  });
}

/** The date column a range applies to: the one given, the lowest date level, else the table's date column. */
function dateColumn(targets: Target[], info: ModelInfo | undefined, given?: string) {
  const table = info?.tables.find((t) => t.name === targets[0]?.table);
  if (!table) return given;
  if (given) return table.columns.find((c) => c.name.toLowerCase() === given.toLowerCase())?.name ?? given;
  const dates = table.columns.filter((c) => /date/i.test(c.type ?? ""));
  const level = [...targets].reverse().find((t) => dates.some((c) => c.name === (t.column ?? t.hierarchyLevel)));
  return level?.column ?? level?.hierarchyLevel ?? (dates.find((c) => /^date$/i.test(c.name)) ?? dates[0])?.name ?? targets[0]?.hierarchy?.replace(/ ?Hierarchy$/i, "");
}

const isoDate = (v: string | number) => {
  const d = new Date(String(v));
  if (Number.isNaN(d.getTime())) throw new Error(`"${v}" is not a date (give it as YYYY-MM-DD).`);
  return `DATE(${d.getUTCFullYear()}, ${d.getUTCMonth() + 1}, ${d.getUTCDate()})`;
};

/** The tree of a hierarchy slicer: each path's last level Selected, its parents Inherited. */
export function tree(paths: unknown[][]): Node[] {
  const root: Node[] = [];
  for (const path of paths) {
    let level = root;
    path.forEach((value, i) => {
      const last = i === path.length - 1;
      let node = level.find((n) => n.value === value);
      if (!node) level.push((node = { value, operator: last ? "Selected" : "Inherited", ...(last ? {} : { children: [] }) }));
      if (!last) level = (node.children ??= []);
    });
  }
  return root;
}

/** The paths of the dates between from and to, one per date, as the slicer's levels show them. */
export async function datePaths(model: Ref, targets: Target[], info: ModelInfo | undefined, range: { from?: string | number; to?: string | number; column?: string }) {
  const table = targets[0].table;
  const date = dateColumn(targets, info, range.column);
  if (!date) throw new Error(`No date column found in '${table}' for this range: give range.column.`);
  const d = col(table, date);
  const cond = [range.from !== undefined && `${d} >= ${isoDate(range.from)}`, range.to !== undefined && `${d} < ${isoDate(range.to)} + 1`].filter(Boolean).join(" && ") || "TRUE()";
  const levels = levelsOf(targets, info);
  const columns = levels.map((l, i) => `"L${i}", ${l.column ? `CALCULATE(MAX(${col(l.table, l.column)}))` : l.part ? l.part(d) : `BLANK()`}`);
  const rows = await dax(model, `EVALUATE DISTINCT(SELECTCOLUMNS(FILTER(ALL(${d}), ${cond}), ${columns.join(", ")}))`);
  if (!rows.length) throw new Error(`No date of '${table}'[${date}] between ${range.from ?? "…"} and ${range.to ?? "…"}: the model has no data there.`);
  return { paths: rows.map((r) => levels.map((_, i) => r[`L${i}`])), date };
}

async function slicerOf(ctx: ToolContext, name: string) {
  const view = (await ctx.browser("report.state")) as ViewState & { slicers: Slicer[] };
  const s = view.slicers.find((x) => x.name === name) ?? view.slicers.find((x) => (x.title ?? "").toLowerCase() === name.toLowerCase());
  if (!s) throw new Error(`No slicer "${name}" on the current page. Slicers: ${view.slicers.map((x) => `"${x.title ?? ""}" (${x.name}) on ${x.fields.map((f) => `${f.table}[${f.column}]`).join(">")}`).join(", ") || "none"}. A slicer may be on another page: go_to_page first, or use set_filter.`);
  return s;
}

export const setSlicer = tool({
  name: "set_slicer",
  description:
    "Select values in a slicer of the current page (by its name or title; find the right slicer by the column it is on, even when its title differs from the user's words). values: exact values of the column (matched without case; a wrong value is refused with the valid ones) — for a hierarchy slicer, values of its top level. paths: for a hierarchy slicer, full paths, e.g. [[2027, \"February\"], [2027, \"March\"]] or [[\"Europe\", \"France\"]]. range: numbers or dates (YYYY-MM-DD), also on a date hierarchy slicer (its years, months and days are selected; range.column: the date column if not obvious). clear: select everything again.",
  args: z.object({
    slicer: z.string(),
    values: z.array(z.union([z.string(), z.number(), z.boolean()])).optional(),
    paths: z.array(z.array(z.union([z.string(), z.number()]))).optional(),
    range: z.object({ from: z.union([z.string(), z.number()]).optional(), to: z.union([z.string(), z.number()]).optional(), column: z.string().optional() }).optional(),
    clear: z.boolean().optional(),
  }),
  changes: true,
  async run(a, ctx) {
    const s = await slicerOf(ctx, a.slicer);
    const report = ctx.scope.report;
    const model = report?.datasetId ? { id: report.datasetId, groupId: report.datasetGroupId } : undefined;
    const targets = (s.targets ?? s.fields.map((f) => ({ table: f.table, column: f.column }))) as Target[];
    const hierarchical = targets.length > 1 || !!targets[0]?.hierarchyLevel;
    const info = model && hierarchical ? await modelInfo(model).catch(() => undefined) : undefined;
    const clear = a.clear || (!a.values?.length && !a.paths?.length && !a.range);
    let args: Record<string, unknown> = { name: s.name, clear };
    let said = "";
    if (!clear && hierarchical) {
      let paths: unknown[][] = a.paths?.map((p) => [...p]) ?? a.values?.map((v) => [v]) ?? [];
      let fallback: unknown;
      if (a.range) {
        if (!model) throw new Error("A range on a hierarchy slicer needs the report's model.");
        const r = await datePaths(model, targets, info, a.range);
        paths = r.paths;
        fallback = { table: targets[0].table, column: r.date, range: a.range };
        said = ` ${paths.length} dates selected through the slicer's levels.`;
      }
      args = { ...args, hierarchy: tree(paths), fallback };
    } else if (!clear) {
      let values = a.values;
      const field = s.fields[0];
      if (values?.length && field && model) values = (await checkValues(model, field.table, field.column, values)) as typeof values;
      args = { ...args, values, range: a.range && { from: a.range.from, to: a.range.to } };
    }
    const done = (await ctx.browser("report.setSlicer", args)) as { note?: string };
    const after = (await ctx.browser("report.state")) as ViewState;
    return `Done.${said}${done?.note ? ` ${done.note}` : ""} ${viewText(after).split("\n").filter((l) => l.includes(`name=${s.name}`) || l.startsWith("Page filters")).join("\n")}`;
  },
});
