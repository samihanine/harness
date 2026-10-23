/** A report's structure read from its PBIR definition: settings, pages and visuals (for the Reports screen). */
import { definition, partText } from "./report-def";
import { embedReport, resetEmbed, type Ref } from "./pbi";
import { all, patch, type ReportEntry } from "./store";

export type VisualSchema = { name: string; page: string; type: string; title?: string; fields: Record<string, string[]>; x: number; y: number; width: number; height: number; hidden: boolean; filters: number };
export type PageSchema = { name: string; displayName: string; width: number; height: number; hidden: boolean; filters: number; visuals: number };
/** `source`: read from the definition (PBIR, everything) or through an embed (read-only reports: no theme nor report measures). */
export type ReportSchema = { source?: "definition" | "embed"; readAt?: string; theme?: string; baseTheme?: string; filters: string[]; settings: Record<string, unknown>; pages: PageSchema[]; visuals: VisualSchema[]; measures: { table: string; name: string; expression: string }[] };

const literal = (v: unknown) => String((v as any)?.expr?.Literal?.Value ?? "").replace(/^'(.*)'$/s, "$1").replace(/''/g, "'");
/** "Table.Column" (or the aggregation) of a query projection. */
const fieldName = (p: any) => p?.queryRef ?? p?.nativeQueryRef ?? "?";
/** The column or measure a filter is on. */
const filterField = (f: any) => {
  const field = f?.field?.Column ?? f?.field?.Measure ?? f?.field?.Aggregation?.Expression?.Column;
  return field ? `${field.Expression?.SourceRef?.Entity ?? "?"}.${field.Property}` : (f?.name ?? "filter");
};

export async function readReportSchema(ref: Ref): Promise<ReportSchema> {
  const parts = await definition(ref);
  const json = (path: string) => {
    const part = parts.find((p) => p.path === path);
    return part ? JSON.parse(partText(part)) : undefined;
  };
  const report = json("definition/report.json") ?? {};
  const order: string[] = json("definition/pages/pages.json")?.pageOrder ?? [];
  const pageFiles = parts.filter((p) => /^definition\/pages\/[^/]+\/page\.json$/.test(p.path));
  const visuals: VisualSchema[] = parts
    .filter((p) => /^definition\/pages\/[^/]+\/visuals\/[^/]+\/visual\.json$/.test(p.path))
    .map((p) => {
      const v = JSON.parse(partText(p));
      const page = p.path.split("/")[2];
      const state = v.visual?.query?.queryState ?? {};
      return {
        name: v.name,
        page,
        type: v.visual?.visualType ?? (v.visualGroup ? "group" : "?"),
        title: literal(v.visual?.visualContainerObjects?.title?.[0]?.properties?.text) || undefined,
        fields: Object.fromEntries(Object.entries<any>(state).map(([role, s]) => [role, (s.projections ?? []).map(fieldName)])),
        x: Math.round(v.position?.x ?? 0),
        y: Math.round(v.position?.y ?? 0),
        width: Math.round(v.position?.width ?? 0),
        height: Math.round(v.position?.height ?? 0),
        hidden: !!v.isHidden,
        filters: v.filterConfig?.filters?.length ?? 0,
      };
    });
  const pages: PageSchema[] = pageFiles
    .map((p) => JSON.parse(partText(p)))
    .map((p) => ({ name: p.name, displayName: p.displayName ?? p.name, width: p.width, height: p.height, hidden: p.visibility === "HiddenInViewMode", filters: p.filterConfig?.filters?.length ?? 0, visuals: visuals.filter((v) => v.page === p.name).length }))
    .sort((a, b) => (order.indexOf(a.name) + 1 || 999) - (order.indexOf(b.name) + 1 || 999));
  const ext = json("definition/reportExtensions.json");
  return {
    theme: report.themeCollection?.customTheme?.name,
    baseTheme: report.themeCollection?.baseTheme?.name,
    filters: (report.filterConfig?.filters ?? []).map(filterField),
    settings: report.settings ?? {},
    pages,
    visuals,
    measures: (ext?.entities ?? []).flatMap((e: any) => (e.measures ?? []).map((m: any) => ({ table: e.name, name: m.name, expression: m.expression }))),
  };
}

const fieldText = (t: Record<string, any> | undefined) => (!t ? "?" : t.measure ? `${t.table}.${t.measure}` : `${t.aggregationFunction ? `${t.aggregationFunction}(` : ""}${t.table}.${t.column ?? t.hierarchyLevel ?? t.hierarchy}${t.aggregationFunction ? ")" : ""}`);

/**
 * The structure of a report the API cannot read (view only, sensitivity label…): what a hidden embed
 * shows — pages, visuals (type, title, place, fields when the SDK gives them), filters.
 */
export async function readReportSchemaEmbedded(ref: Ref): Promise<ReportSchema> {
  const element = Object.assign(document.createElement("div"), { style: "position:fixed;left:-10000px;top:0;width:1280px;height:720px" });
  document.body.append(element);
  try {
    const report = await embedReport(element, ref, "view");
    const pages = await report.getPages();
    const visuals: VisualSchema[] = [];
    const pageSchemas: PageSchema[] = [];
    for (const p of pages) {
      const list = await p.getVisuals().catch(() => []);
      const pageFilters = await p.getFilters().catch(() => []);
      for (const v of list) {
        const fields: Record<string, string[]> = {};
        // Fields: through the authoring SDK when the report allows it (often refused in view mode).
        const caps = await v.getCapabilities().catch(() => undefined);
        for (const role of caps?.dataRoles ?? []) {
          const targets = await v.getDataFields(role.name).catch(() => []);
          if (targets.length) fields[role.name] = targets.map((t) => fieldText(t as never));
        }
        const filters = await v.getFilters().catch(() => []);
        visuals.push({ name: v.name, page: p.name, type: v.type, title: v.title || undefined, fields, x: Math.round(v.layout.x ?? 0), y: Math.round(v.layout.y ?? 0), width: Math.round(v.layout.width ?? 0), height: Math.round(v.layout.height ?? 0), hidden: v.layout.displayState?.mode === 1, filters: filters.length });
      }
      pageSchemas.push({ name: p.name, displayName: p.displayName, width: p.defaultSize?.width ?? 1280, height: p.defaultSize?.height ?? 720, hidden: p.visibility === 1, filters: pageFilters.length, visuals: list.length });
    }
    const reportFilters = await report.getFilters().catch(() => []);
    return { source: "embed", readAt: new Date().toISOString(), filters: reportFilters.map((f: any) => fieldText(f.target)), settings: {}, pages: pageSchemas, visuals, measures: [] };
  } finally {
    await resetEmbed(element).catch(() => undefined);
    element.remove();
  }
}

/**
 * A report's descriptive schema, kept with it (the AI's context, the Reports screen, the Builder's
 * examples): from its definition when readable, else through an embed.
 */
export async function describeReport(entry: ReportEntry): Promise<ReportSchema> {
  const ref = { id: entry.id, groupId: entry.groupId, appId: entry.appId };
  const schema = entry.appId
    ? await readReportSchemaEmbedded(ref)
    : await readReportSchema(ref).then(
        (s) => ({ ...s, source: "definition" as const, readAt: new Date().toISOString() }),
        () => readReportSchemaEmbedded(ref),
      );
  await patch("reports", entry.id, { schema });
  return schema;
}

/** The structure as text for the AI. */
export const schemaText = (s: ReportSchema) =>
  [
    s.source === "embed" ? "Read through an embed (the definition is not readable): no theme nor report measures." : `Theme: ${s.theme ?? s.baseTheme ?? "default"}.`,
    `Report filters: ${s.filters.join(", ") || "none"}.`,
    `Pages (${s.pages.length}):\n${s.pages.map((p) => `- "${p.displayName}" (${p.name}) ${p.width}×${p.height}${p.hidden ? " hidden" : ""}, ${p.visuals} visuals, ${p.filters} filters`).join("\n")}`,
    `Visuals:\n${s.visuals.map((v) => `- [${s.pages.find((p) => p.name === v.page)?.displayName ?? v.page}] ${v.type}${v.title ? ` "${v.title}"` : ""} (${v.name}) x=${v.x} y=${v.y} w=${v.width} h=${v.height}${Object.keys(v.fields).length ? ` ${Object.entries(v.fields).map(([r, f]) => `${r}: ${f.join(", ")}`).join("; ")}` : ""}`).join("\n")}`,
    s.measures.length ? `Report measures:\n${s.measures.map((m) => `- '${m.table}'[${m.name}] = ${m.expression.replace(/\s+/g, " ").slice(0, 160)}`).join("\n")}` : "",
  ]
    .filter(Boolean)
    .join("\n\n");

/** The Builder's reference reports (Reports tab) with their schemas (made when missing). */
export async function referenceReports(ids: string[] = []) {
  if (!ids.length) return [];
  const reports = (await all("reports")).filter((r) => ids.includes(r.id));
  return Promise.all(reports.map(async (r) => ({ name: r.name, schema: r.schema ?? (await describeReport(r).catch(() => undefined)) })));
}

/** Reference reports for the Builder's AI: each one's schema, cut to stay short. */
export async function examplesText(ids: string[] = []) {
  const reports = await referenceReports(ids);
  if (!reports.length) return "";
  const parts = reports.map((r) => (r.schema ? `### Reference report "${r.name}"\n${schemaText(r.schema).slice(0, 4000)}` : `### Reference report "${r.name}": its structure could not be read.`));
  return `Reference reports chosen by the user (Reports tab): context for inspiration, not to change. The user may refer to them ("like …", "the page … of …"); follow their layout, visual types and fields when it fits, and copy_report_page rebuilds one of their pages here.\n${parts.join("\n\n")}`;
}
