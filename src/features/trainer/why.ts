/**
 * "Why did it change?" on a time chart: the app finds the last two periods of the visual (with the
 * filters and slicer selections of its page), breaks the change down by the dimensions the report
 * itself uses on that page (series, slicers, other visuals' axes) with DAX, and the AI only tells
 * what these exact figures say.
 */
import type { VisualDescriptor } from "powerbi-client";
import { AI_MODELS, openSession } from "@/agent/llm";
import { asMarkdown } from "./explain";
import { round } from "@/features/trainer/facts";
import { dax, type Report } from "@/services/pbi";
import { literal } from "@/lib/query-spec";
import type { ModelInfo } from "@/services/model";
import { settings, type DatasetEntry } from "@/lib/store";
import { whyPrompt } from "@/prompts/trainer";

type Field = { table?: string; column?: string; measure?: string; hierarchy?: string; hierarchyLevel?: string; aggregationFunction?: string };

/** Charts with a time axis (role holding the periods, role holding the value). */
const TIME_CHARTS: Record<string, [string, string]> = {
  lineChart: ["Category", "Y"],
  areaChart: ["Category", "Y"],
  stackedAreaChart: ["Category", "Y"],
  columnChart: ["Category", "Y"],
  clusteredColumnChart: ["Category", "Y"],
  lineClusteredColumnComboChart: ["Category", "Y"],
  lineStackedColumnComboChart: ["Category", "Y"],
  ribbonChart: ["Category", "Y"],
  kpi: ["TrendLine", "Indicator"],
};
const TIME_NAME = /date|year|month|week|quarter|day|period|ann[ée]e|mois|semaine|trimestre|jour|p[ée]riode|fiscal|ym\b/i;
const AGG: Record<string, string> = { Sum: "SUM", Average: "AVERAGE", Avg: "AVERAGE", Count: "COUNTA", CountNonNull: "COUNTA", Min: "MIN", Max: "MAX", DistinctCount: "DISTINCTCOUNT" };

const q = (t: string, c: string) => `'${t.replace(/'/g, "''")}'[${c}]`;
const columnOf = (info: ModelInfo | undefined, f: Field) => {
  const name = f.column ?? f.hierarchyLevel;
  if (!f.table || !name) return undefined;
  const c = info?.tables.find((t) => t.name === f.table)?.columns.find((x) => x.name === name);
  return info && !c ? undefined : { table: f.table, name, type: c?.type?.toLowerCase(), sortBy: c?.sortBy };
};
const fields = async (v: VisualDescriptor, role: string) => ((await v.getDataFields(role).catch(() => [])) as Field[]) ?? [];

/** The visual's time column and value, when it is a time chart (else undefined: no button). */
export async function whyTarget(v: VisualDescriptor, info?: ModelInfo) {
  const roles = TIME_CHARTS[v.type];
  if (!roles) return undefined;
  const time = columnOf(info, (await fields(v, roles[0]))[0] ?? {});
  if (!time || !(time.type?.startsWith("date") || TIME_NAME.test(time.name))) return undefined;
  const value = (await fields(v, roles[1]))[0];
  const expr = value?.measure ? `[${value.measure}]` : value?.column && value.table ? `${AGG[value.aggregationFunction ?? "Sum"] ?? "SUM"}(${q(value.table, value.column)})` : undefined;
  if (!expr) return undefined;
  return { time, value: { expr, label: value.measure ?? `${value.aggregationFunction ?? "Sum"} of ${value.column}` } };
}

/** A filter of the SDK as a DAX filter argument (undefined when it cannot be written: then listed as ignored). */
function filterDax(f: any, info?: ModelInfo): string | undefined {
  const c = f?.target && columnOf(info, f.target);
  if (!c) return undefined;
  const col = q(c.table, c.name);
  const lit = (v: unknown) => literal(v, c.type);
  if (Array.isArray(f.values)) {
    if (f.operator === "All" || !f.values.length) return "";
    const set = `{${f.values.map(lit).join(", ")}}`;
    return f.operator === "NotIn" ? `FILTER(ALL(${col}), NOT ${col} IN ${set})` : `TREATAS(${set}, ${col})`;
  }
  if (Array.isArray(f.conditions) && f.conditions.length) {
    const ops: Record<string, string> = { LessThan: "<", LessThanOrEqual: "<=", GreaterThan: ">", GreaterThanOrEqual: ">=", Is: "=", IsNot: "<>" };
    const parts = f.conditions.map((x: any) => (ops[x.operator] ? `${col} ${ops[x.operator]} ${lit(x.value)}` : x.operator === "Contains" ? `CONTAINSSTRING(${col}, ${literal(String(x.value), "string")})` : x.operator === "DoesNotContain" ? `NOT CONTAINSSTRING(${col}, ${literal(String(x.value), "string")})` : undefined));
    if (parts.some((p: unknown) => !p)) return undefined;
    return `FILTER(ALL(${col}), ${parts.join(f.logicalOperator === "Or" ? " || " : " && ")})`;
  }
  return undefined;
}

const active = (filters: any[]) => filters.filter((f) => (f.values?.length && f.operator !== "All") || f.conditions?.length || f.hierarchyData?.length);

export async function whyChanged(report: Report, page: string, visual: string, context: string, dataset: DatasetEntry) {
  const target = (await report.getPages()).find((p) => p.name === page);
  const visuals = (await target?.getVisuals()) ?? [];
  const v = visuals.find((x) => x.name === visual);
  if (!v || !target) throw new Error("Visual not found on the page.");
  const info = dataset.info;
  const spec = await whyTarget(v, info);
  if (!spec) throw new Error("This visual has no time axis with a value: nothing to compare.");
  const ref = { id: dataset.id, groupId: dataset.groupId };

  // Every filter the visual is under: report, page, its own, the slicers of the page.
  const slicers = visuals.filter((x) => x.type === "slicer");
  const all = active([
    ...(await report.getFilters().catch(() => [])),
    ...(await target.getFilters().catch(() => [])),
    ...(await v.getFilters().catch(() => [])),
    ...(await Promise.all(slicers.map(async (s) => ((await s.getSlicerState().catch(() => undefined))?.filters ?? []) as any[]))).flat(),
  ]);
  const written = all.map((f) => ({ f, dax: filterDax(f, info) }));
  const filters = written.map((x) => x.dax).filter((x): x is string => !!x);
  const ignored = written.filter((x) => x.dax === undefined).length;

  // The last two periods with a value, in the column's own order (its sort-by column when it has one).
  const { time, value } = spec;
  const timeCol = q(time.table, time.name);
  const sortCol = time.sortBy ? q(time.table, time.sortBy) : undefined;
  const series = await dax(ref, `EVALUATE SUMMARIZECOLUMNS(${[timeCol, sortCol, ...filters].filter(Boolean).join(", ")}, "v", ${value.expr})\nORDER BY ${sortCol ?? timeCol}`);
  const points = series.filter((r) => r.v !== null && r.v !== undefined && r.v !== "");
  if (points.length < 2) throw new Error("Fewer than two periods with a value under the current filters.");
  const [prev, cur] = points.slice(-2);
  const curV = Number(cur.v);
  const prevV = Number(prev.v);
  const at = (p: Record<string, unknown>) => `CALCULATE(${value.expr}, ${timeCol} = ${literal(p[time.name], time.type)})`;

  // Dimensions the report itself uses on this page: series / legend, slicers, other visuals' axes.
  const dims = new Map<string, { table: string; name: string }>();
  const add = (f: Field) => {
    const c = columnOf(info, f);
    if (c && !(c.table === time.table && c.name === time.name) && !TIME_NAME.test(c.name) && !c.type?.startsWith("date")) dims.set(`${c.table}|${c.name}`, c);
  };
  for (const role of ["Series", "Legend"]) (await fields(v, role)).forEach(add);
  for (const s of slicers) ((await s.getSlicerState().catch(() => undefined))?.targets ?? []).forEach((t) => add(t as Field));
  for (const o of visuals.filter((x) => x !== v && x.type !== "slicer")) for (const role of ["Category", "Rows", "Legend", "Series"]) (await fields(o, role)).forEach(add);
  const chosen = [...dims.values()].slice(0, 6);
  const delta = curV - prevV;
  const breakdowns = await Promise.all(
    chosen.map(async (d) => {
      const rows = await dax(ref, `EVALUATE SUMMARIZECOLUMNS(${[q(d.table, d.name), ...filters].join(", ")}, "cur", ${at(cur)}, "prev", ${at(prev)})`).catch(() => undefined);
      if (!rows) return "";
      const items = rows.map((r) => ({ name: String(r[d.name] ?? "(blank)"), cur: Number(r.cur ?? 0), prev: Number(r.prev ?? 0) })).map((x) => ({ ...x, d: x.cur - x.prev }));
      const top = items.sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 5);
      return `- by ${q(d.table, d.name)}: ${top.map((x) => `${x.name} ${x.d >= 0 ? "+" : ""}${round(x.d)} (${round(x.prev)} → ${round(x.cur)}${delta ? `, ${Math.round((x.d / delta) * 100)} % of the change` : ""})`).join("; ")}`;
    }),
  );
  const facts = [
    `Value: ${value.label}. Periods (${q(time.table, time.name)}): ${String(prev[time.name])} → ${String(cur[time.name])}.`,
    `Change: ${round(prevV)} → ${round(curV)} = ${delta >= 0 ? "+" : ""}${round(delta)}${prevV ? ` (${delta >= 0 ? "+" : ""}${Math.round((delta / Math.abs(prevV)) * 100)} %)` : ""}.`,
    `Earlier periods: ${points.slice(-8, -2).map((p) => `${String(p[time.name])} ${round(Number(p.v))}`).join(", ") || "none"}.`,
    breakdowns.filter(Boolean).length ? `Biggest contributions to the change (items moving the most):\n${breakdowns.filter(Boolean).join("\n")}` : "No dimension of the page to break the change down.",
    ignored ? `${ignored} filter(s) of the page could not be applied to these figures (relative dates, hierarchies…): they may differ slightly from the visual.` : "",
  ].filter(Boolean);
  const { model } = await settings.get();
  const prompt = whyPrompt(context, facts);
  return asMarkdown(await openSession(model ?? AI_MODELS[0], "why").send(prompt));
}
