/**
 * A page of a reference report (its schema: types, titles, places, fields — read from the definition
 * or, when it is closed, through an embed) as visual specs for the edited report: the fields checked
 * against this report's model, the places scaled to its page. What cannot be rebuilt is listed.
 */
import { resolveField, type ModelInfo } from "@/services/model";
import type { ReportSchema } from "@/services/report-schema";
import { ROLES, type ReportMeasures, type VisualSpec } from "./visual-spec";

const AGG: Record<string, string> = { sum: "Sum", average: "Avg", avg: "Avg", count: "Count", countnonnull: "CountNonNull", min: "Min", max: "Max" };

/** "Table.Name" (the table may hold dots): the split whose table and field exist. */
function split(ref: string, info: ModelInfo | undefined, extension: ReportMeasures) {
  const dots = [...ref.matchAll(/\./g)].map((m) => m.index!);
  for (const i of dots) {
    const table = ref.slice(0, i);
    const name = ref.slice(i + 1);
    if (extension.some((m) => m.table === table && m.name === name)) return `[${name}]`;
    try {
      const f = resolveField(info, `'${table}'[${name}]`);
      return `'${f.table}'[${f.name}]`;
    } catch {}
  }
  return undefined;
}

/** A field of a schema ("Sales.Amount", "Sum(Sales.Amount)", "Sales.Total Sales") as a field spec, or undefined when this model lacks it. */
export function fieldSpec(ref: string, info: ModelInfo | undefined, extension: ReportMeasures) {
  const m = ref.trim().match(/^(\w+)\((.+)\)$/);
  if (m) {
    const fn = AGG[m[1].toLowerCase()];
    const inner = split(m[2], info, extension);
    return fn && inner ? `${fn}(${inner})` : undefined;
  }
  return split(ref.trim(), info, extension);
}

/** Roles of the definition as create_visual names them (a chart's "Rows" are its small multiples). */
const roleOf = (type: string, role: string) => (role === "Rows" && !ROLES[type]?.includes("Rows") && ROLES[type]?.includes("SmallMultiples") ? "SmallMultiples" : role);

export function pageCopy(schema: ReportSchema, page: string, info: ModelInfo | undefined, extension: ReportMeasures, target: { width: number; height: number }) {
  const p = schema.pages.find((x) => x.name === page || x.displayName.toLowerCase() === page.toLowerCase());
  if (!p) throw new Error(`No page "${page}" in that report. Pages: ${schema.pages.map((x) => `"${x.displayName}"`).join(", ")}.`);
  const sx = target.width / (p.width || 1280);
  const sy = target.height / (p.height || 720);
  const visuals: (VisualSpec & { source: string })[] = [];
  const dropped: string[] = [];
  /** Report measures of the reference report the copied visuals use (made in this report when missing). */
  const measures = new Map<string, ReportSchema["measures"][number]>();
  for (const v of schema.visuals.filter((x) => x.page === p.name && !x.hidden)) {
    const label = `${v.type}${v.title ? ` "${v.title}"` : ""}`;
    if (!ROLES[v.type] || v.type === "textbox" || v.type === "actionButton") {
      dropped.push(`${label}: not rebuilt (${v.type === "textbox" ? "its text is not in the schema" : "type not built here"})`);
      continue;
    }
    const fields: Record<string, string[]> = {};
    const missing: string[] = [];
    for (const [role0, list] of Object.entries(v.fields)) {
      const role = roleOf(v.type, role0);
      if (!ROLES[v.type].includes(role)) {
        missing.push(...list.map((f) => `${f} (role ${role0})`));
        continue;
      }
      for (const f of list) {
        const own = schema.measures.find((m) => `${m.table}.${m.name}` === f);
        const known = fieldSpec(f, info, [...extension, ...(own ? [{ table: own.table, name: own.name }] : [])]);
        if (own && known) measures.set(own.name, own);
        if (known) (fields[role] ??= []).push(known);
        else missing.push(f);
      }
    }
    const hadFields = Object.values(v.fields).some((l) => l.length);
    if (hadFields && !Object.keys(fields).length) {
      dropped.push(`${label}: none of its fields is in this model (${missing.join(", ")})`);
      continue;
    }
    if (missing.length) dropped.push(`${label}: without ${missing.join(", ")} (not in this model)`);
    visuals.push({ source: v.name, type: v.type, title: v.title, fields, x: Math.round(v.x * sx), y: Math.round(v.y * sy), width: Math.round(v.width * sx), height: Math.round(v.height * sy) });
  }
  return { page: p, visuals, dropped, measures: [...measures.values()] };
}
