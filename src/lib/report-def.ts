/**
 * Report definition (PBIR) through Fabric: report-level measures ("extension" measures, stored in
 * definition/reportExtensions.json). They only need write access to the report, not to the model.
 */
import { fabric } from "./ms";
import { dax, type Ref } from "./pbi";
import type { ModelInfo } from "./model";
import { workspaceId } from "./tmdl";

type Part = { path: string; payload: string; payloadType: string };
type Extension = { $schema: string; name: string; entities?: { name: string; measures?: ExtensionMeasure[] }[] };
export type ExtensionMeasure = {
  name: string;
  dataType: string;
  expression: string;
  formatString?: string;
  displayFolder?: string;
  description?: string;
  references?: { measures?: { entity: string; name: string; schema?: string }[]; unrecognizedReferences?: boolean };
};

const PATH = "definition/reportExtensions.json";
const SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/reportExtension/1.0.0/schema.json";
const decode = (p: string) => new TextDecoder().decode(Uint8Array.from(atob(p), (c) => c.charCodeAt(0)));
const encode = (s: string) => btoa(String.fromCharCode(...new TextEncoder().encode(s)));

export async function definition(report: Ref): Promise<Part[]> {
  const ws = await workspaceId(report.groupId);
  const result = await fabric(`/workspaces/${ws}/reports/${report.id}/getDefinition?format=PBIR`, "POST").catch((e) => {
    throw new Error(`The report definition is not readable in PBIR format (${e instanceof Error ? e.message : e}). Save the report once from the editor, and check you can edit it.`);
  });
  return result.definition.parts.filter((p: Part) => p.path !== ".platform");
}

export async function readExtension(report: Ref): Promise<Extension> {
  const part = (await definition(report)).find((p) => p.path === PATH);
  return part ? JSON.parse(decode(part.payload)) : { $schema: SCHEMA, name: "extension", entities: [] };
}

/** Report measures as text for the AI. */
export async function reportMeasuresText(report: Ref) {
  const ext = await readExtension(report);
  const list = (ext.entities ?? []).flatMap((e) => (e.measures ?? []).map((m) => `[${m.name}] (in '${e.name}', report measure) = ${m.expression}`));
  return list.length ? list.join("\n") : "No report measure yet.";
}

const DATA_TYPE = (v: unknown) =>
  typeof v === "boolean" ? "Boolean" : typeof v === "number" ? (Number.isInteger(v) ? "Integer" : "Double") : typeof v === "string" && /^\d{4}-\d\d-\d\dT/.test(v) ? "DateTime" : typeof v === "string" ? "Text" : "Variant";

/**
 * Adds (or replaces) a report measure: the DAX is first evaluated on the model (errors come back as
 * they are), then written to the report definition.
 */
export async function upsertReportMeasure(
  report: Ref,
  dataset: Ref,
  model: ModelInfo | undefined,
  m: { table: string; name: string; expression: string; format?: string; folder?: string; description?: string },
) {
  const parts = await definition(report);
  const part = parts.find((p) => p.path === PATH);
  const ext: Extension = part ? JSON.parse(decode(part.payload)) : { $schema: SCHEMA, name: "extension", entities: [] };
  const others = (ext.entities ?? []).flatMap((e) => (e.measures ?? []).filter((x) => x.name !== m.name).map((x) => ({ entity: e.name, x })));
  // The other report measures are defined too, so the new one can use them.
  const defs = [...others.map(({ entity, x }) => `MEASURE '${entity}'[${x.name}] = ${x.expression}`), `MEASURE '${m.table}'[${m.name}] = ${m.expression}`];
  const [row] = await dax(dataset, `DEFINE\n${defs.join("\n")}\nEVALUATE ROW("value", [${m.name}])`);
  const modelMeasures = model?.tables.flatMap((t) => t.measures.map((x) => ({ entity: t.name, name: x.name }))) ?? [];
  const used = [...m.expression.matchAll(/\[([^\]]+)\]/g)].map((x) => x[1]);
  const references = [
    ...modelMeasures.filter((x) => used.includes(x.name)).map((x) => ({ ...x, schema: "" })),
    ...others.filter(({ x }) => used.includes(x.name)).map(({ entity, x }) => ({ entity, name: x.name, schema: ext.name })),
  ];
  const measure: ExtensionMeasure = {
    name: m.name,
    dataType: DATA_TYPE(row?.value),
    expression: m.expression,
    ...(m.format && { formatString: m.format }),
    ...(m.folder && { displayFolder: m.folder }),
    ...(m.description && { description: m.description }),
    ...(references.length && { references: { measures: references, unrecognizedReferences: false } }),
  };
  const entities = (ext.entities ?? []).map((e) => ({ ...e, measures: (e.measures ?? []).filter((x) => x.name !== m.name) })).filter((e) => e.measures.length || e.name === m.table);
  const entity = entities.find((e) => e.name === m.table);
  if (entity) entity.measures.push(measure);
  else entities.push({ name: m.table, measures: [measure] });
  const next = { ...ext, $schema: ext.$schema ?? SCHEMA, name: ext.name ?? "extension", entities };
  const nextParts = [...parts.filter((p) => p.path !== PATH), { path: PATH, payload: encode(JSON.stringify(next, null, 2)), payloadType: "InlineBase64" }];
  await fabric(`/workspaces/${await workspaceId(report.groupId)}/reports/${report.id}/updateDefinition`, "POST", { definition: { parts: nextParts } });
  return `Report measure [${m.name}] saved in '${m.table}' (value now: ${row?.value ?? "blank"}).`;
}

export async function deleteReportMeasure(report: Ref, name: string) {
  const parts = await definition(report);
  const part = parts.find((p) => p.path === PATH);
  if (!part) return "No report measure.";
  const ext: Extension = JSON.parse(decode(part.payload));
  ext.entities = (ext.entities ?? []).map((e) => ({ ...e, measures: (e.measures ?? []).filter((x) => x.name !== name) })).filter((e) => e.measures.length);
  await fabric(`/workspaces/${await workspaceId(report.groupId)}/reports/${report.id}/updateDefinition`, "POST", {
    definition: { parts: [...parts.filter((p) => p.path !== PATH), { path: PATH, payload: encode(JSON.stringify(ext, null, 2)), payloadType: "InlineBase64" }] },
  });
  return `Report measure [${name}] deleted.`;
}


/** Edits JSON files of the report definition (PBIR), then writes it back. `edit` returns a summary. */
async function editJson(report: Ref, edit: (read: (path: string) => any, write: (path: string, value: unknown) => void, paths: string[]) => string) {
  const parts = await definition(report);
  const changed = new Map<string, string>();
  const read = (path: string) => {
    const p = parts.find((x) => x.path === path);
    if (!p && !changed.has(path)) throw new Error(`Not found in the report definition: ${path}`);
    return JSON.parse(changed.get(path) ?? decode(p!.payload));
  };
  const summary = edit(read, (path, value) => changed.set(path, JSON.stringify(value, null, 2)), parts.map((p) => p.path));
  // Written paths that did not exist are new files (e.g. a new visual).
  const added = [...changed.keys()].filter((path) => !parts.some((p) => p.path === path));
  const next = [
    ...parts.map((p) => (changed.has(p.path) ? { ...p, payload: encode(changed.get(p.path)!) } : p)),
    ...added.map((path) => ({ path, payload: encode(changed.get(path)!), payloadType: "InlineBase64" })),
  ];
  await fabric(`/workspaces/${await workspaceId(report.groupId)}/reports/${report.id}/updateDefinition`, "POST", { definition: { parts: next } });
  return summary;
}

/** Text of a text box (the authoring API cannot write it): one paragraph per line. */
export function setTextbox(report: Ref, page: string, visual: string, t: { text: string; size?: number; bold?: boolean; color?: string; align?: "left" | "center" | "right" }) {
  return editJson(report, (read, write) => {
    const path = `definition/pages/${page}/visuals/${visual}/visual.json`;
    const json = read(path);
    if (json.visual?.visualType !== "textbox") throw new Error(`${visual} is a ${json.visual?.visualType}, not a text box.`);
    const textStyle = { ...(t.bold && { fontWeight: "bold" }), ...(t.size && { fontSize: `${t.size}pt` }), ...(t.color && { color: t.color }) };
    const paragraphs = t.text.split("\n").map((line) => ({ ...(t.align && { horizontalTextAlignment: t.align }), textRuns: [{ value: line, textStyle }] }));
    json.visual.objects = { ...json.visual.objects, general: [{ properties: { paragraphs } }] };
    json.isHidden = undefined;
    write(path, json);
    return `Text box ${visual}: "${t.text.slice(0, 60)}"${t.size ? ` ${t.size}pt` : ""}${t.bold ? " bold" : ""}.`;
  });
}

type Box = { x?: number; y?: number; width?: number; height?: number };
const visualPaths = (paths: string[], page: string) => paths.filter((p) => p.startsWith(`definition/pages/${page}/visuals/`) && p.endsWith("/visual.json"));

/** Positions / sizes of several visuals at once, written in the definition (reliable, one reload). */
export function setPositions(report: Ref, page: string, moves: (Box & { visual: string })[]) {
  return editJson(report, (read, write) => {
    for (const m of moves) {
      const path = `definition/pages/${page}/visuals/${m.visual}/visual.json`;
      const json = read(path);
      const box = Object.fromEntries((["x", "y", "width", "height"] as const).filter((k) => m[k] !== undefined).map((k) => [k, Math.round(m[k]!)]));
      json.position = { ...json.position, ...box };
      write(path, json);
    }
    return `Moved ${moves.length} visual(s).`;
  });
}

/**
 * A colored shape (rectangle by default). `behind`: placed under every other visual of the page,
 * e.g. a colored panel grouping visuals.
 */
export function addShape(report: Ref, page: string, s: Required<Box> & { color: string; shape?: "rectangle" | "roundedRectangle" | "oval"; behind?: boolean; transparency?: number }) {
  return editJson(report, (read, write, paths) => {
    const others = visualPaths(paths, page).map((path) => ({ path, json: read(path) }));
    const zs = others.map((o) => o.json.position?.z ?? 0);
    let z = zs.length ? Math.max(...zs) + 1000 : 0;
    if (s.behind) {
      // Everything else moves one step to the front, the shape takes the lowest place.
      z = zs.length ? Math.min(...zs) : 0;
      for (const o of others) write(o.path, { ...o.json, position: { ...o.json.position, z: (o.json.position?.z ?? 0) + 1 } });
    }
    const name = Array.from(crypto.getRandomValues(new Uint8Array(10)), (b) => b.toString(16).padStart(2, "0")).join("");
    const lit = (v: string) => ({ expr: { Literal: { Value: v } } });
    write(`definition/pages/${page}/visuals/${name}/visual.json`, {
      $schema: others[0]?.json.$schema ?? "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json",
      name,
      position: { x: Math.round(s.x), y: Math.round(s.y), z, width: Math.round(s.width), height: Math.round(s.height), tabOrder: z },
      visual: {
        visualType: "shape",
        objects: {
          shape: [{ properties: { tileShape: lit(`'${s.shape === "roundedRectangle" ? "rectangleRounded" : s.shape === "oval" ? "oval" : "rectangle"}'`), ...(s.shape === "roundedRectangle" && { roundEdge: lit("12L") }) } }],
          fill: [{ properties: { show: lit("true"), fillColor: { solid: { color: lit(`'${s.color}'`) } }, transparency: lit(`${s.transparency ?? 0}D`) }, selector: { id: "default" } }],
          outline: [{ properties: { show: lit("false") } }],
        },
        drawVisualDisplayBorder: false,
      },
    });
    return `Shape ${name} (${s.color}) x=${Math.round(s.x)} y=${Math.round(s.y)} w=${Math.round(s.width)} h=${Math.round(s.height)}${s.behind ? ", behind the other visuals" : ""}.`;
  });
}

export type Interaction = "filter" | "highlight" | "none" | "default";
const INTERACTION = { filter: "DataFilter", highlight: "HighlightFilter", none: "NoFilter", default: "Default" } as const;

/** How a click / selection in `source` acts on each target visual of the page ("Edit interactions"). */
export function setInteractions(report: Ref, page: string, source: string, targets: { visual: string; type: Interaction }[]) {
  return editJson(report, (read, write) => {
    const path = `definition/pages/${page}/page.json`;
    const json = read(path);
    const others = (json.visualInteractions ?? []).filter((i: { source: string; target: string }) => !(i.source === source && targets.some((t) => t.visual === i.target)));
    json.visualInteractions = [...others, ...targets.filter((t) => t.type !== "default").map((t) => ({ source, target: t.visual, type: INTERACTION[t.type] }))];
    write(path, json);
    return `Interactions of ${source} set: ${targets.map((t) => `${t.visual} → ${t.type}`).join(", ")}.`;
  });
}

/** Single or multiple selection of a slicer (and the "Select all" option). */
export function setSlicerSelection(report: Ref, page: string, visual: string, mode: { single?: boolean; selectAll?: boolean; style?: "dropdown" | "list" | "tile" }) {
  return editJson(report, (read, write) => {
    const path = `definition/pages/${page}/visuals/${visual}/visual.json`;
    const json = read(path);
    if (json.visual?.visualType !== "slicer") throw new Error(`${visual} is a ${json.visual?.visualType}, not a slicer.`);
    const lit = (v: boolean) => ({ expr: { Literal: { Value: String(v) } } });
    const text = (v: string) => ({ expr: { Literal: { Value: `'${v}'` } } });
    json.visual.objects = {
      ...json.visual.objects,
      ...((mode.single !== undefined || mode.selectAll !== undefined) && {
        selection: [
          {
            properties: {
              ...json.visual.objects?.selection?.[0]?.properties,
              ...(mode.single !== undefined && { singleSelect: lit(mode.single), strictSingleSelect: lit(mode.single) }),
              ...(mode.selectAll !== undefined && { selectAllCheckboxEnabled: lit(mode.selectAll) }),
            },
          },
        ],
      }),
      // Slicer style: "Dropdown" (a select), "Basic" (a list) or "Tile" (buttons).
      ...(mode.style && { data: [{ properties: { mode: text(mode.style === "dropdown" ? "Dropdown" : mode.style === "tile" ? "Tile" : "Basic") } }] }),
    };
    write(path, json);
    return `Slicer ${visual}:${mode.single === undefined ? "" : ` ${mode.single ? "single" : "multiple"} selection`}${mode.style ? ` ${mode.style} style` : ""}.`;
  });
}

/** PBIR expressions in short form: literals, colors, fields ('T'[c], Sum('T'[c]), [Measure]). */
function short(v: any): any {
  if (v === null || typeof v !== "object") return v;
  if (Array.isArray(v)) return v.map(short);
  if (v.expr) return short(v.expr);
  if (v.Literal) return String(v.Literal.Value).replace(/^'(.*)'$/s, "$1").replace(/^(-?[\d.]+)[DLM]$/, "$1");
  if (v.solid) return short(v.solid.color);
  if (v.ThemeDataColor) return `theme color ${v.ThemeDataColor.ColorId}${v.ThemeDataColor.Percent ? ` ${Math.round(v.ThemeDataColor.Percent * 100)}%` : ""}`;
  const entity = (x: any) => x?.Expression?.SourceRef?.Entity ?? x?.Expression?.SourceRef?.Source;
  if (v.Column) return `'${entity(v.Column)}'[${v.Column.Property}]`;
  if (v.Measure) return `'${entity(v.Measure)}'[${v.Measure.Property}]`;
  if (v.HierarchyLevel) return `'${entity(v.HierarchyLevel.Expression?.Hierarchy)}'[${v.HierarchyLevel.Expression?.Hierarchy?.Hierarchy}].[${v.HierarchyLevel.Level}]`;
  if (v.Aggregation) return `${["Sum", "Avg", "Count", "Min", "Max", "CountNonNull", "Median", "StdDev", "Var"][v.Aggregation.Function] ?? "Agg"}(${short(v.Aggregation.Expression)})`;
  return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, short(x)]));
}
const compact = (v: unknown) => JSON.stringify(short(v));

/** Format objects ({title: [{properties, selector}]}) as one line each: `title: text="Sales" show=true`. */
function objectsText(objects: Record<string, any[]> | undefined, indent: string) {
  return Object.entries(objects ?? {}).flatMap(([name, entries]) =>
    (entries ?? []).map((e) => {
      const props = Object.entries(short(e.properties ?? {})).map(([k, x]) => `${k}=${typeof x === "string" ? JSON.stringify(x) : JSON.stringify(x)}`);
      return `${indent}${name}${e.selector ? ` [${compact(e.selector)}]` : ""}: ${props.join(" ")}`;
    }),
  );
}

/**
 * Everything needed to rebuild a report, from its definition (PBIR): theme, pages (size, background,
 * interactions, filters), visuals (type, position, fields, formatting, filters, text), report
 * measures, bookmarks. Parameters and model measures live in the semantic model (see its summary).
 */
export function definitionText(parts: Part[]) {
  const json = (path: string) => {
    const p = parts.find((x) => x.path === path);
    try { return p ? JSON.parse(decode(p.payload)) : undefined; } catch { return undefined; }
  };
  const out: string[] = ["Report definition (PBIR): enough to rebuild it."];
  const report = json("definition/report.json");
  if (report) {
    const theme = report.themeCollection;
    out.push(`theme: base ${theme?.baseTheme?.name ?? "default"}${theme?.customTheme ? `, custom ${theme.customTheme.name}` : ""}`);
    out.push(...objectsText(report.objects, "report "));
    if (report.filterConfig?.filters?.length) out.push(`report filters: ${compact(report.filterConfig.filters)}`);
  }
  // Custom theme (colors, default visual styles).
  for (const p of parts.filter((x) => /RegisteredResources\/.+\.json$/.test(x.path))) {
    const t = json(p.path);
    if (t?.dataColors || t?.visualStyles || t?.background) out.push(`custom theme ${p.path.split("/").pop()}: ${JSON.stringify(t).slice(0, 4000)}`);
  }
  const order: string[] = json("definition/pages/pages.json")?.pageOrder ?? [];
  const pages = [...new Set([...order, ...parts.map((p) => p.path.match(/^definition\/pages\/([^/]+)\/page\.json$/)?.[1]).filter(Boolean) as string[]])];
  for (const name of pages) {
    const page = json(`definition/pages/${name}/page.json`);
    if (!page) continue;
    out.push("", `## Page "${page.displayName}" (name=${name}) ${page.width ?? 1280}×${page.height ?? 720}${page.visibility === "HiddenInViewMode" ? " hidden" : ""}`);
    out.push(...objectsText(page.objects, "page "));
    if (page.filterConfig?.filters?.length) out.push(`page filters: ${compact(page.filterConfig.filters)}`);
    for (const i of page.visualInteractions ?? []) out.push(`interaction: ${i.source} → ${i.target}: ${i.type}`);
    const visuals = parts
      .map((p) => p.path.match(new RegExp(`^definition/pages/${name}/visuals/([^/]+)/visual\\.json$`))?.[1])
      .filter(Boolean)
      .map((v) => json(`definition/pages/${name}/visuals/${v}/visual.json`))
      .filter(Boolean)
      .sort((a, b) => (a.position?.z ?? 0) - (b.position?.z ?? 0));
    for (const v of visuals) {
      const { x = 0, y = 0, z = 0, width = 0, height = 0 } = v.position ?? {};
      const type = v.visual?.visualType ?? (v.visualGroup ? "group" : "?");
      out.push(`- ${type} name=${v.name} x=${Math.round(x)} y=${Math.round(y)} w=${Math.round(width)} h=${Math.round(height)} z=${z}${v.parentGroupName ? ` in group ${v.parentGroupName}` : ""}${v.isHidden ? " (hidden)" : ""}`);
      if (v.visualGroup) out.push(`  group "${v.visualGroup.displayName ?? ""}"`);
      const roles = v.visual?.query?.queryState ?? {};
      for (const [role, state] of Object.entries<any>(roles)) {
        const fields = (state.projections ?? []).map((p: any) => `${short(p.field)}${p.displayName ? ` as "${p.displayName}"` : ""}`);
        if (fields.length) out.push(`  ${role}: ${fields.join(", ")}`);
      }
      if (v.visual?.query?.sortDefinition?.sort?.length) out.push(`  sort: ${v.visual.query.sortDefinition.sort.map((s: any) => `${short(s.field)} ${s.direction}`).join(", ")}`);
      out.push(...objectsText(v.visual?.objects, "  format "));
      out.push(...objectsText(v.visual?.visualContainerObjects, "  container "));
      if (v.filterConfig?.filters?.length) out.push(`  filters: ${compact(v.filterConfig.filters)}`);
    }
  }
  const ext = json(PATH) as Extension | undefined;
  const measures = (ext?.entities ?? []).flatMap((e) => (e.measures ?? []).map((m) => `[${m.name}] (in '${e.name}'${m.formatString ? `, format ${m.formatString}` : ""}) = ${m.expression}`));
  if (measures.length) out.push("", "Report measures:", ...measures);
  const bookmarks = parts.filter((p) => /^definition\/bookmarks\/.+\.bookmark\.json$/.test(p.path)).map((p) => json(p.path)).filter(Boolean);
  if (bookmarks.length) out.push("", "Bookmarks:", ...bookmarks.map((b: any) => `- "${b.displayName}" (page ${b.explorationState?.activeSection ?? "?"})`));
  return out.join("\n");
}

/**
 * The authoring API binds report measures as if they were model measures ("this field was deleted
 * from the model"): their references need Schema "extension" and the measure's home table.
 */
export function fixExtensionRefs(report: Ref, page: string, visual: string, measures: { table: string; name: string }[]) {
  return editJson(report, (read, write) => {
    const path = `definition/pages/${page}/visuals/${visual}/visual.json`;
    const json = read(path);
    let fixed = 0;
    const walk = (v: any) => {
      if (!v || typeof v !== "object") return;
      // Bound as a column (the measure was unknown when the visual was made): it becomes a measure.
      if (v.Column && measures.some((x) => x.name === v.Column.Property)) {
        v.Measure = v.Column;
        delete v.Column;
      }
      const ref = v.Measure?.Expression?.SourceRef;
      const m = ref && measures.find((x) => x.name === v.Measure.Property);
      if (m && (ref.Schema !== "extension" || ref.Entity !== m.table)) {
        v.Measure.Expression.SourceRef = { Schema: "extension", Entity: m.table };
        fixed++;
      }
      for (const [k, x] of Object.entries(v)) {
        walk(x);
        // An aggregation of what is now a measure (e.g. Count of the former column): the measure itself.
        const inner = (x as any)?.Aggregation?.Expression;
        if (inner?.Measure?.Expression?.SourceRef?.Schema === "extension") v[k] = inner;
      }
    };
    walk(json);
    if (fixed) write(path, json);
    return `${fixed} report measure reference(s) fixed in ${visual}.`;
  });
}
