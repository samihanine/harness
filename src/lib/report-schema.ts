/** A report's structure read from its PBIR definition: settings, pages and visuals (for the Reports screen). */
import { definition, partText } from "./report-def";
import type { Ref } from "./pbi";

export type VisualSchema = { name: string; page: string; type: string; title?: string; fields: Record<string, string[]>; x: number; y: number; width: number; height: number; hidden: boolean; filters: number };
export type PageSchema = { name: string; displayName: string; width: number; height: number; hidden: boolean; filters: number; visuals: number };
export type ReportSchema = { theme?: string; baseTheme?: string; filters: string[]; settings: Record<string, unknown>; pages: PageSchema[]; visuals: VisualSchema[]; measures: { table: string; name: string; expression: string }[] };

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
