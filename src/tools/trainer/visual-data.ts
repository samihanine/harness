/**
 * The data a visual shows now (with the current filters), as the visual summarizes it. Power BI's export of
 * a visual's data can be refused (tenant setting, visual type): the same figures are then queried with DAX
 * from the visual's fields and the filters acting on it (report, page, visual, slicers).
 */
import { z } from "zod";
import { dax, rowsText } from "@server/pbi/dax";
import { tool } from "../tool";

type Target = { table: string; column?: string; measure?: string; hierarchyLevel?: string; aggregationFunction?: string };
type Filter = { target?: Target | Target[]; operator?: string; values?: unknown[]; conditions?: { operator: string; value: unknown }[]; logicalOperator?: string; hierarchyData?: unknown };

const col = (t: Target) => `'${t.table.replace(/'/g, "''")}'[${(t.column ?? t.hierarchyLevel ?? "").replace(/]/g, "]]")}]`;
const lit = (v: unknown) => (typeof v === "number" || typeof v === "boolean" ? String(v).toUpperCase() : /^\d{4}-\d{2}-\d{2}T/.test(String(v)) ? `DATEVALUE("${String(v).slice(0, 10)}")` : `"${String(v).replace(/"/g, '""')}"`);
const OPS: Record<string, string> = { LessThan: "<", LessThanOrEqual: "<=", GreaterThan: ">", GreaterThanOrEqual: ">=", Is: "=", IsNot: "<>" };
const AGG: Record<string, string> = { Sum: "SUM", Average: "AVERAGE", Avg: "AVERAGE", Count: "COUNTA", CountNonNull: "COUNTA", Min: "MIN", Max: "MAX", DistinctCount: "DISTINCTCOUNT" };

/** A DAX filter table for a Power BI filter (basic In / advanced conditions); undefined when not expressible. */
function filterTable(f: Filter) {
  const t = Array.isArray(f.target) ? undefined : f.target;
  if (!t?.table || !(t.column || t.hierarchyLevel)) return undefined;
  if (f.values?.length && (f.operator ?? "In") === "In") return `TREATAS({${f.values.map(lit).join(", ")}}, ${col(t)})`;
  if (f.conditions?.length) {
    const parts = f.conditions.map((c) => (OPS[c.operator] ? `${col(t)} ${OPS[c.operator]} ${lit(c.value)}` : c.operator === "Contains" ? `CONTAINSSTRING(${col(t)}, ${lit(c.value)})` : undefined));
    if (parts.some((p) => !p)) return undefined;
    return `FILTER(ALL(${col(t)}), ${parts.join(f.logicalOperator === "Or" ? " || " : " && ")})`;
  }
  return undefined;
}

export function visualDax(fields: Record<string, Target[]>, filters: Filter[], rows: number) {
  const all = Object.values(fields).flat();
  const groups = [...new Set(all.filter((t) => (t.column || t.hierarchyLevel) && !t.aggregationFunction).map(col))];
  const values = all.filter((t) => t.measure || t.aggregationFunction).map((t, i) => (t.measure ? `"${t.measure}", [${t.measure}]` : `"${t.aggregationFunction} of ${t.column}_${i}", ${AGG[t.aggregationFunction!] ?? "SUM"}(${col(t)})`));
  const tables = filters.map(filterTable);
  const skipped = tables.filter((x) => !x).length;
  const args = [...groups, ...tables.filter(Boolean), ...[...new Set(values)]];
  if (!values.length && !groups.length) throw new Error("This visual has no field to query.");
  return { query: `EVALUATE TOPN(${rows}, SUMMARIZECOLUMNS(${args.join(", ")}))`, skipped };
}

export const visualData = tool({
  name: "visual_data",
  description: "The data a visual of the current page shows now, with the current filters (its summarized rows, as CSV). Use it to read exact figures of a visual.",
  args: z.object({ visual: z.string().describe("Name or title of the visual."), rows: z.number().optional().describe("Default 50, max 500.") }),
  async run(a, ctx) {
    const rows = Math.min(500, a.rows ?? 50);
    try {
      const r = (await ctx.browser("report.visualData", { visual: a.visual, rows })) as { title?: string; csv: string };
      return `Visual "${r.title ?? a.visual}":\n${r.csv.trim() || "(no data)"}`;
    } catch (e) {
      const why = e instanceof Error ? e.message : String(e);
      const report = ctx.scope.report;
      if (!report?.datasetId) throw e;
      const q = (await ctx.browser("report.visualQuery", { visual: a.visual })) as { title?: string; fields: Record<string, Target[]>; filters: Filter[] };
      const { query, skipped } = visualDax(q.fields, q.filters, rows);
      const data = await dax({ id: report.datasetId, groupId: report.datasetGroupId }, query);
      return `Visual "${q.title ?? a.visual}" (Power BI refused to export its data: ${why}; the same figures queried with DAX from its fields and the current filters${skipped ? `, ${skipped} filter(s) not expressible were left out` : ""}):\n${rowsText(data, rows)}\nQuery: ${query}`;
    }
  },
});
