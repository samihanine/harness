/**
 * Reads the structure of a semantic model with DAX only (executeQueries), so it works with
 * Build / Read permission. Each part is tried on its own and falls back to what is allowed:
 * INFO.VIEW.* → COLUMNSTATISTICS() → measures seen in reports. Failures become warnings.
 */
import type { DatasetRef } from "./api";
import { executeDax } from "./api";
import type { ModelContent, TableInfo } from "./types";

const text = (v: unknown) => (v === null || v === undefined || v === "" ? undefined : String(v));
const message = (e: unknown) => (e instanceof Error ? e.message : String(e)).slice(0, 200);

export async function readModel(ref: DatasetRef, knownMeasures: string[] = []): Promise<ModelContent> {
  const warnings: string[] = [];
  const methods: string[] = [];
  const query = async (name: string, dax = `EVALUATE ${name}()`) => {
    try {
      const { rows } = await executeDax(ref, dax);
      methods.push(name);
      return rows;
    } catch (error) {
      warnings.push(`${name}: ${message(error)}`);
      return null;
    }
  };

  const [tables, columns, measures, relationships] = await Promise.all([
    query("INFO.VIEW.TABLES"),
    query("INFO.VIEW.COLUMNS"),
    query("INFO.VIEW.MEASURES"),
    query("INFO.VIEW.RELATIONSHIPS"),
  ]);
  // Scans the model: only when columns are not listable otherwise.
  const stats = columns ? null : await query("COLUMNSTATISTICS");

  const byName = new Map<string, TableInfo>();
  const table = (name: string) => {
    if (!byName.has(name)) byName.set(name, { name, columns: [], measures: [] });
    return byName.get(name)!;
  };

  for (const t of tables ?? []) Object.assign(table(String(t.Name)), { description: text(t.Description), hidden: t.IsHidden === true || undefined });
  for (const c of columns ?? []) {
    if (String(c.Name).startsWith("RowNumber-")) continue;
    table(String(c.Table)).columns.push({
      name: String(c.Name),
      dataType: text(c.DataType),
      description: text(c.Description),
      displayFolder: text(c.DisplayFolder),
      hidden: c.IsHidden === true || undefined,
    });
  }
  for (const m of measures ?? [])
    table(String(m.Table)).measures.push({
      name: String(m.Name),
      expression: text(m.Expression),
      formatString: text(m.FormatString),
      description: text(m.Description),
      displayFolder: text(m.DisplayFolder),
      hidden: m.IsHidden === true || undefined,
    });

  // COLUMNSTATISTICS: columns when INFO.VIEW is not allowed (with value ranges).
  for (const s of stats ?? []) {
    const name = String(s["Column Name"] ?? "");
    if (!name || name.startsWith("RowNumber-")) continue;
    const t = table(String(s["Table Name"]));
    let column = t.columns.find((c) => c.name === name);
    if (!column) t.columns.push((column = { name, dataType: inferType(s.Min, s.Max) }));
    column.stats = { min: s.Min ?? undefined, max: s.Max ?? undefined, cardinality: Number(s.Cardinality) || undefined };
  }

  // Measures used by reports, when the model does not list them.
  if (!measures && knownMeasures.length) {
    const t = table("(measures seen in reports)");
    for (const name of new Set(knownMeasures)) t.measures.push({ name });
    warnings.push("Measure definitions are not readable: names come from the reports using this model.");
  }

  if (byName.size === 0) throw new Error(`Nothing readable in this semantic model.\n${warnings.join("\n")}`);
  return {
    tables: [...byName.values()].filter((t) => t.columns.length || t.measures.length),
    relationships: (relationships ?? []).map((r) => ({
      from: `'${String(r.FromTable)}'[${String(r.FromColumn)}]`,
      to: `'${String(r.ToTable)}'[${String(r.ToColumn)}]`,
      active: r.IsActive !== false,
      crossFilter: text(r.CrossFilteringBehavior),
    })),
    warnings,
    methods,
    extractedAt: new Date().toISOString(),
  };
}

function inferType(min: unknown, max: unknown) {
  const sample = min ?? max;
  if (typeof sample === "number") return Number.isInteger(sample) ? "Int64" : "Double";
  if (typeof sample === "boolean") return "Boolean";
  if (typeof sample === "string" && /^\d{4}-\d\d-\d\dT/.test(sample)) return "DateTime";
  return "String";
}

/** Compact text of a model for the agent (one line per table). */
export function modelSummary(model: ModelContent, { expressions = true } = {}) {
  // Power BI auto date tables and tables with nothing visible are noise for the agent.
  const tables = model.tables.filter(
    (t) => !/^(LocalDateTable|DateTableTemplate)_/.test(t.name) && (t.measures.length || t.columns.some((c) => !c.hidden)),
  );
  const lines = tables.map((t) => {
    const cols = t.columns.filter((c) => !c.hidden).map((c) => `${c.name}${c.dataType ? `:${c.dataType}` : ""}`);
    const ms = t.measures.map((m) => (expressions && m.expression ? `[${m.name}] = ${m.expression.replace(/\s+/g, " ").slice(0, 300)}` : `[${m.name}]`));
    return `'${t.name}'${t.description ? ` (${t.description})` : ""}\n  columns: ${cols.join(", ") || "-"}${ms.length ? `\n  measures: ${ms.join(" | ")}` : ""}`;
  });
  const rels = model.relationships.map((r) => `${r.from} → ${r.to}${r.active ? "" : " (inactive)"}`);
  return [...lines, rels.length ? `relationships:\n  ${rels.join("\n  ")}` : "", model.warnings.length ? `limits: ${model.warnings.join("; ")}` : ""]
    .filter(Boolean)
    .join("\n");
}
