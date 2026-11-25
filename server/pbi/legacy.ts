/**
 * Reports saved in the old format (PBIR-Legacy: one report.json with sections and visual containers whose
 * config is a JSON string) shown to the Builder as PBIR files, so every draft tool works on them; on save,
 * only the pages and visuals that changed are written back into report.json — everything else is kept as is.
 * Queries: PBIR's queryState (fields with their table) ⇄ the legacy projections + prototypeQuery (fields
 * referring to aliases of a From list).
 */
import { partText, textPart, type Part } from "./definitions";

type Json = Record<string, any>;
export type Legacy = { report: Json; pbir: Part; originals: Map<string, Json> };

const DISPLAY = ["FitToPage", "FitToPage", "FitToWidth", "ActualSize"];
const parse = (s: unknown, fallback: any = {}) => {
  if (typeof s !== "string" || !s) return fallback;
  try {
    return JSON.parse(s);
  } catch {
    return fallback;
  }
};
const clone = <T>(x: T): T => (x === undefined ? x : JSON.parse(JSON.stringify(x)));

export const isLegacy = (parts: Part[]) => parts.some((p) => p.path === "report.json") && !parts.some((p) => p.path.startsWith("definition/"));

/* ------------------------------ Fields ⇄ aliases ------------------------------ */

/** Replaces every SourceRef of a field: { Source: alias } → { Entity, Schema } (legacy → PBIR). */
function withEntities(x: any, from: Record<string, { Entity: string; Schema?: string }>): any {
  if (Array.isArray(x)) return x.map((y) => withEntities(y, from));
  if (!x || typeof x !== "object") return x;
  if (x.SourceRef?.Source && from[x.SourceRef.Source]) return { ...withEntities({ ...x, SourceRef: undefined }, from), SourceRef: { ...(from[x.SourceRef.Source].Schema ? { Schema: from[x.SourceRef.Source].Schema } : {}), Entity: from[x.SourceRef.Source].Entity } };
  return Object.fromEntries(Object.entries(x).filter(([, v]) => v !== undefined).map(([k, v]) => [k, withEntities(v, from)]));
}

/** The other way: { Entity, Schema } → { Source: alias }, the aliases collected in `from`. */
function withAliases(x: any, from: Json[]): any {
  if (Array.isArray(x)) return x.map((y) => withAliases(y, from));
  if (!x || typeof x !== "object") return x;
  if (x.SourceRef?.Entity) {
    const { Entity, Schema } = x.SourceRef;
    let hit = from.find((f) => f.Entity === Entity && f.Schema === Schema);
    if (!hit) {
      const base = Entity.replace(/[^a-z]/gi, "").slice(0, 1).toLowerCase() || "t";
      let name = base;
      for (let i = 1; from.some((f) => f.Name === name); i++) name = `${base}${i}`;
      hit = { Name: name, Entity, ...(Schema ? { Schema } : {}), Type: 0 };
      from.push(hit);
    }
    return { ...Object.fromEntries(Object.entries(x).filter(([k]) => k !== "SourceRef").map(([k, v]) => [k, withAliases(v, from)])), SourceRef: { Source: hit.Name } };
  }
  return Object.fromEntries(Object.entries(x).map(([k, v]) => [k, withAliases(v, from)]));
}

/** A legacy visual filter (expression) as PBIR writes it (field), and back. */
const filterToPbir = (f: Json) => {
  const { expression, ...rest } = f;
  return { ...rest, ...(expression ? { field: expression } : {}) };
};
const filterToLegacy = (f: Json) => {
  const { field, ...rest } = f;
  return { ...rest, ...(field ? { expression: field } : {}) };
};

/* ----------------------------- Legacy → PBIR files ----------------------------- */

function visualToPbir(container: Json): Json {
  const config = parse(container.config);
  const pos = config.layouts?.[0]?.position ?? { x: container.x, y: container.y, z: container.z, width: container.width, height: container.height };
  const out: Json = { name: config.name, position: { x: pos.x ?? 0, y: pos.y ?? 0, z: pos.z ?? 0, width: pos.width ?? 300, height: pos.height ?? 200, tabOrder: pos.tabOrder ?? pos.z ?? 0 } };
  const single = config.singleVisual;
  if (single) {
    const from = Object.fromEntries((single.prototypeQuery?.From ?? []).map((f: Json) => [f.Name, { Entity: f.Entity, Schema: f.Schema }]));
    const selects = new Map<string, Json>((single.prototypeQuery?.Select ?? []).map((s: Json) => [s.Name, s]));
    const queryState: Json = {};
    for (const [role, list] of Object.entries<any[]>(single.projections ?? {})) {
      queryState[role] = {
        projections: list.map((p) => {
          const s = selects.get(p.queryRef) ?? {};
          const { Name, NativeReferenceName, ...field } = s;
          const display = single.columnProperties?.[p.queryRef]?.displayName;
          return { field: withEntities(field, from), queryRef: p.queryRef, nativeQueryRef: NativeReferenceName ?? String(p.queryRef).split(".").at(-1), ...(p.active !== undefined ? { active: p.active } : {}), ...(display ? { displayName: display } : {}) };
        }),
      };
    }
    const order = single.prototypeQuery?.OrderBy;
    out.visual = {
      visualType: single.visualType,
      ...(Object.keys(queryState).length ? { query: { queryState, ...(order?.length ? { sortDefinition: { sort: order.map((o: Json) => ({ field: withEntities(o.Expression, from), direction: o.Direction === 2 ? "Descending" : "Ascending" })) } } : {}) } } : {}),
      ...(single.objects ? { objects: single.objects } : {}),
      ...(single.vcObjects ? { visualContainerObjects: single.vcObjects } : {}),
      drillFilterOtherVisuals: single.drillFilterOtherVisuals ?? true,
    };
    if (single.display?.mode === "hidden") out.isHidden = true;
  } else if (config.singleVisualGroup) out.visualGroup = { displayName: config.singleVisualGroup.displayName, groupMode: config.singleVisualGroup.groupMode ?? "ScaleMode" };
  if (config.parentGroupName) out.parentGroupName = config.parentGroupName;
  const filters = parse(container.filters, []);
  if (filters.length) out.filterConfig = { filters: filters.map(filterToPbir) };
  return out;
}

function pageToPbir(section: Json): Json {
  const config = parse(section.config);
  const filters = parse(section.filters, []);
  return {
    name: section.name,
    displayName: section.displayName,
    displayOption: DISPLAY[section.displayOption ?? 1] ?? "FitToPage",
    width: section.width ?? 1280,
    height: section.height ?? 720,
    ...(config.objects ? { objects: config.objects } : {}),
    ...(filters.length ? { filterConfig: { filters: filters.map(filterToPbir) } } : {}),
    ...(config.visibility === 1 ? { visibility: "HiddenInViewMode" } : {}),
  };
}

/** The legacy definition as PBIR files (the draft works on these), and what is needed to write it back. */
export function legacyToPbir(parts: Part[]): { parts: Part[]; legacy: Legacy } {
  const report = JSON.parse(partText(parts.find((p) => p.path === "report.json")!));
  const config = parse(report.config);
  const originals = new Map<string, Json>();
  const out: Part[] = [];
  const json = (path: string, value: unknown) => out.push(textPart(path, JSON.stringify(value, null, 2)));
  const sections = [...(report.sections ?? [])].sort((a, b) => (a.ordinal ?? 0) - (b.ordinal ?? 0));
  json("definition/report.json", { themeCollection: config.themeCollection, ...(config.objects ? { objects: config.objects } : {}), ...(config.settings ? { settings: config.settings } : {}), filterConfig: { filters: parse(report.filters, []).map(filterToPbir) } });
  json("definition/pages/pages.json", { pageOrder: sections.map((s) => s.name), activePageName: sections[config.activeSectionIndex ?? 0]?.name });
  const ext = config.modelExtensions?.[0];
  if (ext?.entities?.length) json("definition/reportExtensions.json", { name: ext.name ?? "extension", entities: ext.entities });
  for (const s of sections) {
    originals.set(`page:${s.name}`, s);
    json(`definition/pages/${s.name}/page.json`, pageToPbir(s));
    for (const c of s.visualContainers ?? []) {
      const v = visualToPbir(c);
      if (!v.name) continue;
      originals.set(`visual:${s.name}/${v.name}`, c);
      json(`definition/pages/${s.name}/visuals/${v.name}/visual.json`, v);
    }
  }
  // Pictures and themes registered in the report: the same paths in both formats.
  out.push(...parts.filter((p) => p.path.startsWith("StaticResources/")));
  const pbir = parts.find((p) => p.path === "definition.pbir")!;
  return { parts: out, legacy: { report, pbir, originals } };
}

/* ----------------------------- PBIR files → legacy ----------------------------- */

function visualToLegacy(v: Json, original?: Json): Json {
  const before = original ? parse(original.config) : {};
  const pos = v.position ?? {};
  const config: Json = { ...before, name: v.name, layouts: [{ id: 0, position: { x: pos.x, y: pos.y, z: pos.z ?? 0, width: pos.width, height: pos.height, tabOrder: pos.tabOrder ?? pos.z ?? 0 } }, ...(before.layouts ?? []).filter((l: Json) => l.id !== 0)] };
  if (v.parentGroupName) config.parentGroupName = v.parentGroupName;
  else delete config.parentGroupName;
  const container: Json = { ...(original ?? {}), x: pos.x, y: pos.y, z: pos.z ?? 0, width: pos.width, height: pos.height };
  if (v.visualGroup) config.singleVisualGroup = { ...(before.singleVisualGroup ?? {}), displayName: v.visualGroup.displayName, groupMode: v.visualGroup.groupMode };
  if (v.visual) {
    const single: Json = { ...(before.singleVisual ?? {}), visualType: v.visual.visualType, drillFilterOtherVisuals: v.visual.drillFilterOtherVisuals ?? true };
    for (const k of ["objects", "vcObjects"]) delete single[k];
    if (v.visual.objects) single.objects = v.visual.objects;
    if (v.visual.visualContainerObjects) single.vcObjects = v.visual.visualContainerObjects;
    if (v.isHidden) single.display = { mode: "hidden" };
    else delete single.display;
    const queryState = v.visual.query?.queryState;
    const oldState = original ? visualToPbir(original).visual?.query : undefined;
    // The query is written again only when it changed (the cached query of the container then goes).
    if (JSON.stringify(v.visual.query ?? null) !== JSON.stringify(oldState ?? null)) {
      delete container.query;
      delete container.dataTransforms;
      delete single.columnProperties;
      const from: Json[] = [];
      const select: Json[] = [];
      const projections: Json = {};
      const columnProperties: Json = {};
      for (const [role, state] of Object.entries<any>(queryState ?? {})) {
        projections[role] = (state.projections ?? []).map((p: Json) => {
          if (!select.some((s) => s.Name === p.queryRef)) select.push({ ...withAliases(p.field, from), Name: p.queryRef, NativeReferenceName: p.nativeQueryRef });
          if (p.displayName) columnProperties[p.queryRef] = { displayName: p.displayName };
          return { queryRef: p.queryRef, ...(p.active !== undefined ? { active: p.active } : {}) };
        });
      }
      const orderBy = (v.visual.query?.sortDefinition?.sort ?? []).map((s: Json) => ({ Direction: s.direction === "Descending" ? 2 : 1, Expression: withAliases(s.field, from) }));
      single.projections = projections;
      single.prototypeQuery = { Version: 2, From: from, Select: select, ...(orderBy.length ? { OrderBy: orderBy } : {}) };
      if (Object.keys(columnProperties).length) single.columnProperties = columnProperties;
      if (!select.length) (delete single.projections, delete single.prototypeQuery);
    }
    config.singleVisual = single;
    if (v.visual.visualType === "actionButton") config.howCreated ??= "InsertVisualButton";
  }
  container.config = JSON.stringify(config);
  container.filters = JSON.stringify((v.filterConfig?.filters ?? []).map(filterToLegacy));
  return container;
}

/** report.json with the draft's changes (`changed`: paths written or removed since the definition was read). */
export function pbirToLegacy(files: Part[], legacy: Legacy, changed: Set<string>): Part[] {
  const read = (path: string) => {
    const p = files.find((x) => x.path === path);
    return p ? JSON.parse(partText(p)) : undefined;
  };
  const touched = (path: string) => [...changed].some((c) => path === c || path.startsWith(`${c}/`) || c.startsWith(`${path}/`) || c === path);
  const report = clone(legacy.report);
  const config = parse(report.config);
  const order: string[] = read("definition/pages/pages.json")?.pageOrder ?? [];
  const pageNames = files.map((p) => p.path.match(/^definition\/pages\/([^/]+)\/page\.json$/)?.[1]).filter(Boolean) as string[];
  const names = [...order.filter((n) => pageNames.includes(n)), ...pageNames.filter((n) => !order.includes(n))];
  report.sections = names.map((name, ordinal) => {
    const original = legacy.originals.get(`page:${name}`);
    const page = read(`definition/pages/${name}/page.json`);
    const section: Json = original ? clone(original) : { name, config: "{}", filters: "[]", visualContainers: [] };
    // The order is the sections' order; an ordinal only where the report writes them.
    if (!original || original.ordinal !== undefined) section.ordinal = ordinal;
    if (!original || touched(`definition/pages/${name}/page.json`)) {
      const sc = parse(section.config);
      section.displayName = page.displayName;
      section.width = page.width;
      section.height = page.height;
      section.displayOption = Math.max(1, DISPLAY.indexOf(page.displayOption ?? "FitToPage"));
      if (page.objects) sc.objects = page.objects;
      if (page.visibility === "HiddenInViewMode") sc.visibility = 1;
      else delete sc.visibility;
      section.config = JSON.stringify(sc);
      section.filters = JSON.stringify((page.filterConfig?.filters ?? []).map(filterToLegacy));
    }
    const visualPaths = files.filter((p) => p.path.startsWith(`definition/pages/${name}/visuals/`) && p.path.endsWith("/visual.json"));
    section.visualContainers = visualPaths.map((p) => {
      const v = JSON.parse(partText(p));
      const key = `visual:${name}/${v.name}`;
      const kept = legacy.originals.get(key);
      return kept && !touched(p.path) ? kept : visualToLegacy(v, kept);
    });
    return section;
  });
  // Report level: theme, settings, objects, filters, report measures.
  const r = read("definition/report.json") ?? {};
  if (touched("definition/report.json")) {
    if (r.themeCollection) config.themeCollection = r.themeCollection;
    if (r.objects) config.objects = r.objects;
    if (r.settings) config.settings = r.settings;
    report.filters = JSON.stringify((r.filterConfig?.filters ?? []).map(filterToLegacy));
  }
  const ext = read("definition/reportExtensions.json");
  if (ext && touched("definition/reportExtensions.json")) config.modelExtensions = [{ ...(config.modelExtensions?.[0] ?? {}), name: ext.name ?? "extension", entities: ext.entities }];
  const active = read("definition/pages/pages.json")?.activePageName;
  config.activeSectionIndex = Math.max(0, names.indexOf(active));
  report.config = JSON.stringify(config);
  return [legacy.pbir, textPart("report.json", JSON.stringify(report)), ...files.filter((p) => p.path.startsWith("StaticResources/"))];
}
