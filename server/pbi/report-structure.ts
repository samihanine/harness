/**
 * The structure of a report, as complete as the user's rights allow:
 * - editable (its PBIR definition is readable): pages, visuals with their fields, sort, filters, positions,
 *   formatting, the theme, report measures, bookmarks — everything;
 * - view only: what a hidden embed shows (pages, visuals, fields of slicers, filters).
 * Used by the AI (inspiration, reading), and downloaded as JSON from the Reports screen.
 */
import type { Ref } from "@shared/types";
import { pbi, scope } from "../ms/api";
import { embeddedStructure } from "../capture/browser";
import { partText, reportDefinition, type Part } from "./definitions";

const literal = (v: unknown) => String((v as any)?.expr?.Literal?.Value ?? "").replace(/^'(.*)'$/s, "$1").replace(/''/g, "'");
const fieldOf = (f: any): string => {
  if (!f) return "?";
  if (f.Measure) return `${f.Measure.Expression?.SourceRef?.Entity ?? "?"}[${f.Measure.Property}]`;
  if (f.Column) return `${f.Column.Expression?.SourceRef?.Entity ?? "?"}[${f.Column.Property}]`;
  if (f.Aggregation) return `${["Sum", "Avg", "Count", "Min", "Max", "CountNonNull"][f.Aggregation.Function] ?? "Agg"}(${fieldOf(f.Aggregation.Expression)})`;
  if (f.HierarchyLevel) return `${f.HierarchyLevel.Expression?.Hierarchy?.Expression?.SourceRef?.Entity ?? "?"}[${f.HierarchyLevel.Level}]`;
  return JSON.stringify(f).slice(0, 80);
};
const filtersOf = (config: any) =>
  (config?.filters ?? []).map((f: any) => ({
    field: fieldOf(f.field),
    type: f.type,
    ...(f.filter?.Where ? { condition: JSON.stringify(f.filter.Where).replace(/"(Expression|SourceRef|Source|Property)":/g, "").slice(0, 300) } : {}),
  }));

export type ReportStructure = { source: "definition" | "embed"; report: { id: string; name?: string }; [k: string]: unknown };

export async function reportStructureFromParts(parts: Part[], detail: "summary" | "full") {
  const json = (path: string) => {
    const p = parts.find((x) => x.path === path);
    return p ? JSON.parse(partText(p)) : undefined;
  };
  const report = json("definition/report.json") ?? {};
  const order: string[] = json("definition/pages/pages.json")?.pageOrder ?? [];
  const pages = parts
    .filter((p) => /^definition\/pages\/[^/]+\/page\.json$/.test(p.path))
    .map((p) => JSON.parse(partText(p)))
    .sort((a, b) => (order.indexOf(a.name) + 1 || 999) - (order.indexOf(b.name) + 1 || 999))
    .map((page) => {
      const visuals = parts
        .filter((p) => p.path.startsWith(`definition/pages/${page.name}/visuals/`) && p.path.endsWith("/visual.json"))
        .map((p) => {
          const v = JSON.parse(partText(p));
          const state = v.visual?.query?.queryState ?? {};
          return {
            name: v.name,
            type: v.visual?.visualType ?? (v.visualGroup ? "group" : "?"),
            title: literal(v.visual?.visualContainerObjects?.title?.[0]?.properties?.text) || undefined,
            position: v.position && { x: Math.round(v.position.x), y: Math.round(v.position.y), width: Math.round(v.position.width), height: Math.round(v.position.height), z: v.position.z },
            fields: Object.fromEntries(Object.entries<any>(state).map(([role, s]) => [role, (s.projections ?? []).map((pr: any) => fieldOf(pr.field))])),
            ...(v.visual?.query?.sortDefinition ? { sort: v.visual.query.sortDefinition.sort?.map((s: any) => `${fieldOf(s.field)} ${s.direction}`) } : {}),
            ...(v.filterConfig?.filters?.length ? { filters: filtersOf(v.filterConfig) } : {}),
            ...(v.isHidden ? { hidden: true } : {}),
            ...(v.parentGroupName ? { group: v.parentGroupName } : {}),
            ...(detail === "full" ? { formatting: { objects: v.visual?.objects, container: v.visual?.visualContainerObjects } } : {}),
          };
        });
      return {
        name: page.name,
        displayName: page.displayName,
        size: { width: page.width, height: page.height },
        ...(page.visibility === "HiddenInViewMode" ? { hidden: true } : {}),
        ...(page.filterConfig?.filters?.length ? { filters: filtersOf(page.filterConfig) } : {}),
        ...(detail === "full" && page.objects ? { formatting: page.objects } : {}),
        visuals,
      };
    });
  const ext = json("definition/reportExtensions.json");
  const themePart = parts.find((p) => /StaticResources\/RegisteredResources\/.+\.json$/.test(p.path) && /theme/i.test(partText(p).slice(0, 4000)));
  const bookmarks = parts.filter((p) => /bookmarks\/[^/]+\.bookmark\.json$/.test(p.path)).map((p) => JSON.parse(partText(p))).map((b) => ({ name: b.name, displayName: b.displayName }));
  return {
    theme: { base: report.themeCollection?.baseTheme?.name, custom: report.themeCollection?.customTheme?.name, ...(detail === "full" && themePart ? { definition: JSON.parse(partText(themePart)) } : {}) },
    reportFilters: filtersOf(report.filterConfig),
    ...(detail === "full" ? { settings: report.settings, objects: report.objects } : {}),
    measures: (ext?.entities ?? []).flatMap((e: any) => (e.measures ?? []).map((m: any) => ({ table: e.name, name: m.name, expression: m.expression, format: m.formatString }))),
    ...(bookmarks.length ? { bookmarks } : {}),
    pages,
  };
}

/** Summaries the browser already holds (kept with the report, refreshed on demand). */
const seeded = new Map<string, ReportStructure>();
export const seedStructure = (id: string, s: unknown) => void (s && typeof s === "object" && Array.isArray((s as { pages?: unknown }).pages) && seeded.set(id, s as ReportStructure));

export async function reportStructure(ref: Ref, detail: "summary" | "full" = "summary", fresh = false): Promise<ReportStructure> {
  if (detail === "summary" && !fresh && seeded.has(ref.id)) return seeded.get(ref.id)!;
  const meta = ref.appId ? undefined : await pbi(`${scope(ref.groupId)}/reports/${ref.id}`).catch(() => undefined);
  const head = { id: ref.id, name: meta?.name, datasetId: meta?.datasetId, webUrl: meta?.webUrl };
  if (!ref.appId) {
    try {
      const parts = await reportDefinition(ref);
      return { source: "definition", report: head, ...(await reportStructureFromParts(parts, detail)) };
    } catch {
      // Not editable: read through an embed.
    }
  }
  return { source: "embed", report: head, note: "View-only report: read through an embed (no formatting, theme nor report measures).", ...(await embeddedStructure(ref)) };
}
