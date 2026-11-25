/**
 * New visuals and pages of a report definition (PBIR) from a short spec the AI fills (type, place,
 * fields by role, title, sort): the JSON is written by this code, the fields checked against the model
 * and the report measures first, so a weak model never has to write PBIR queries by hand.
 */
import catalog from "./format-catalog.json";
import { resolveField, type ModelInfo } from "./model";

/** Data roles of the visual types the app builds (others can still be changed through their JSON). */
export const ROLES: Record<string, string[]> = {
  tableEx: ["Values"],
  pivotTable: ["Rows", "Columns", "Values"],
  card: ["Values"],
  cardVisual: ["Data"],
  multiRowCard: ["Values"],
  lineChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  areaChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  clusteredBarChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  clusteredColumnChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  barChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  columnChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  hundredPercentStackedBarChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  hundredPercentStackedColumnChart: ["Category", "Y", "Series", "SmallMultiples", "Tooltips"],
  lineClusteredColumnComboChart: ["Category", "Y", "Y2", "Series", "SmallMultiples", "Tooltips"],
  pieChart: ["Category", "Y", "Tooltips"],
  donutChart: ["Category", "Y", "Tooltips"],
  slicer: ["Values"],
  kpi: ["Indicator", "TrendLine", "Goal"],
  gauge: ["Y", "MinValue", "MaxValue", "TargetValue"],
  textbox: [],
  shape: [],
  image: [],
  actionButton: [],
};
export const TYPE_NOTES =
  "barChart = stacked bars, columnChart = stacked columns, clustered* = side by side, hundredPercentStacked* = 100 %; SmallMultiples = one small chart per value of a column (grid: format smallMultiplesLayout); slicer with several columns in Values = a hierarchy slicer (Zone > SubZone); textbox = free text (text, text_style); shape = a rectangle, line, oval… (shape, fill: a frame, a band, a separator); image = a picture from a public link (image_url); actionButton = a button with a text and/or a Power BI icon (text, icon, action)";

/** The icons of Power BI buttons (actionButton "icon"). */
export const BUTTON_ICONS = ["blank", "leftArrow", "rightArrow", "reset", "back", "information", "help", "qna", "bookmark", "navigator"];
/** A role as PBIR names it (the small multiples are the "Rows" of a chart). */
export const roleKey = (role: string) => (role === "SmallMultiples" ? "Rows" : role);

const AGGREGATIONS = ["sum", "avg", "count", "min", "max", "countnonnull"];
const lit = (value: string) => ({ expr: { Literal: { Value: value } } });
export const textLiteral = (s: string) => lit(`'${s.replace(/'/g, "''")}'`);

/** Report measures (reportExtensions.json): "Table[Measure]" → home table. */
export type ReportMeasures = { table: string; name: string }[];

/**
 * A field written 'Table'[Name], Table[Name], [Name], Name or Sum(Table[Column]) → PBIR projection.
 * Report measures are referenced with Schema "extension".
 */
export function projection(spec: string, info: ModelInfo | undefined, extension: ReportMeasures) {
  const m = spec.trim().match(/^(\w+)\((.+)\)$/);
  const aggregation = m && AGGREGATIONS.indexOf(m[1].toLowerCase());
  if (m && aggregation! < 0) throw new Error(`"${m[1]}" is not an aggregation (${AGGREGATIONS.join(", ")}).`);
  const inner = m ? m[2] : spec;
  const bare = inner.trim().match(/^\[([^\]]+)\]$/)?.[1] ?? inner.trim().match(/^'?([^'[]+)'?\[([^\]]+)\]$/)?.[2] ?? inner.trim();
  const own = extension.find((x) => x.name.toLowerCase() === bare.toLowerCase());
  if (own && !m) {
    return {
      field: { Measure: { Expression: { SourceRef: { Schema: "extension", Entity: own.table } }, Property: own.name } },
      queryRef: `${own.table}.${own.name}`,
      nativeQueryRef: own.name,
    };
  }
  const known = resolveField(info, inner);
  const ref = { Expression: { SourceRef: { Entity: known.table } }, Property: known.name };
  if (known.kind === "measure") {
    if (m) throw new Error(`[${known.name}] is a measure: use it without ${m[1]}().`);
    return { field: { Measure: ref }, queryRef: `${known.table}.${known.name}`, nativeQueryRef: known.name };
  }
  if (m) {
    const fn = AGGREGATIONS[aggregation!];
    const label = fn === "countnonnull" ? "Count" : fn[0].toUpperCase() + fn.slice(1);
    return { field: { Aggregation: { Expression: { Column: ref }, Function: aggregation } }, queryRef: `${label}(${known.table}.${known.name})`, nativeQueryRef: `${label === "Avg" ? "Average" : label} of ${known.name}` };
  }
  return { field: { Column: ref }, queryRef: `${known.table}.${known.name}`, nativeQueryRef: known.name };
}

/** Visuals showing one value: a bare column shows nothing there, so it is summarized as Power BI does. */
const SINGLE_VALUE = ["card", "cardVisual", "gauge", "kpi"];

/** A column as one value: its first value (text, date, true/false) or its sum (numbers). */
function oneValue(p: ReturnType<typeof projection>, info: ModelInfo | undefined) {
  const column = (p.field as any).Column;
  if (!column) return p;
  const table = column.Expression.SourceRef.Entity;
  const type = info?.tables.find((t) => t.name === table)?.columns.find((c) => c.name === column.Property)?.type ?? "string";
  const numeric = ["int64", "double", "decimal"].includes(type.toLowerCase());
  return {
    field: { Aggregation: { Expression: { Column: column }, Function: numeric ? 0 : 3 } },
    queryRef: `${numeric ? "Sum" : "Min"}(${table}.${column.Property})`,
    nativeQueryRef: `${numeric ? "Sum of" : "First"} ${column.Property}`,
  };
}

export type VisualSpec = {
  type: string;
  x: number;
  y: number;
  width: number;
  height: number;
  fields?: Record<string, string[]>;
  title?: string;
  /** A field of the visual (as given in fields) and its direction. */
  sort?: { field: string; descending?: boolean };
  /** textbox: its text (one paragraph per line). */
  text?: string;
  /** Keep the first N items of the category (or rows) by a value, e.g. a top 10 products by sales. */
  top?: { count: number; by: string; bottom?: boolean };
  /** textbox and button: how the text looks. */
  textStyle?: { fontSize?: number; bold?: boolean; italic?: boolean; color?: string; fontFamily?: string; align?: "left" | "center" | "right" };
  /** image: a public link to the picture. */
  imageUrl?: string;
  /** actionButton: its Power BI icon and what a click does. */
  icon?: string;
  action?: { type: "page" | "bookmark" | "back" | "url"; target?: string };
  /** shape: its form (rectangle, rectangleRounded, line, oval…) and fill color. */
  shape?: string;
  fill?: string;
};

/** The roles holding the items a Top N keeps (the first column found). */
const ITEM_ROLES = ["Category", "Rows", "Values", "Series"];

/**
 * A visual-level Top N filter (as Power BI writes it): the visual's first category column, limited to
 * the `count` first items by a measure or an aggregated column.
 */
export function topNFilter(queryState: Record<string, { projections: { field: any }[] }>, top: NonNullable<VisualSpec["top"]>, info: ModelInfo | undefined, extension: ReportMeasures) {
  const item = ITEM_ROLES.flatMap((r) => queryState[r]?.projections ?? []).find((p) => p.field?.Column)?.field.Column;
  if (!item) throw new Error("A Top N needs a column in the visual (Category or Rows) to keep the first items of.");
  if (!(top.count >= 1)) throw new Error("top.count must be 1 or more.");
  const by = projection(top.by, info, extension).field as any;
  const entity = item.Expression.SourceRef.Entity;
  // The ordering value, referring to its table by a source name of the subquery.
  const byEntity = (by.Measure ?? by.Column ?? by.Aggregation?.Expression?.Column)?.Expression?.SourceRef;
  const withSource = (field: any, source: string) => JSON.parse(JSON.stringify(field).replace(/"SourceRef":\{[^}]*\}/, `"SourceRef":{"Source":"${source}"}`));
  const from = [{ Name: "i", Entity: entity, Type: 0 }, ...(byEntity?.Entity && byEntity.Entity !== entity ? [{ Name: "o", Entity: byEntity.Entity, Type: 0, ...(byEntity.Schema && { Schema: byEntity.Schema }) }] : [])];
  return {
    name: newName(),
    field: { Column: item },
    type: "TopN",
    filter: {
      Version: 2,
      From: [
        { Name: "subquery", Expression: { Subquery: { Query: { Version: 2, From: from, Select: [{ Column: { Expression: { SourceRef: { Source: "i" } }, Property: item.Property }, Name: "field" }], OrderBy: [{ Direction: top.bottom ? 1 : 2, Expression: withSource(by, from.length > 1 ? "o" : "i") }], Top: Math.round(top.count) } } }, Type: 2 },
        { Name: "i", Entity: entity, Type: 0 },
      ],
      Where: [{ Condition: { In: { Expressions: [{ Column: { Expression: { SourceRef: { Source: "i" } }, Property: item.Property } }], Table: { SourceRef: { Source: "subquery" } } } } }],
    },
  };
}

/** A new name (id) for a page or a visual, like Power BI's (20 hex characters). */
export const newName = () => Array.from(crypto.getRandomValues(new Uint8Array(10)), (b) => b.toString(16).padStart(2, "0")).join("");

const VISUAL_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json";
const PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json";

/** Role names models use for another one ("Values" on a cardVisual, whose role is "Data"; "Axis", "Legend" of the old editor). */
const ROLE_ALIASES: Record<string, string[]> = { Values: ["Data", "Y"], Data: ["Values"], Value: ["Values", "Data", "Y"], Y: ["Values"], Axis: ["Category", "Rows"], Legend: ["Series", "Columns"], Category: ["Rows"], Rows: ["Category"] };

/** The fields with their roles as the visual type names them: a wrong role goes to its usual equivalent, or to the type's only role. */
export function fitRoles(type: string, fields: Record<string, string[]> | undefined) {
  const roles = ROLES[type];
  if (!roles || !fields) return fields;
  return Object.fromEntries(
    Object.entries(fields).map(([role, list]) => {
      if (roles.includes(role)) return [role, list];
      const to = ROLE_ALIASES[role]?.find((r) => roles.includes(r) && !(r in fields)) ?? (roles.length === 1 ? roles[0] : undefined);
      return [to ?? role, list];
    }),
  );
}

const colorOf = (c: string) => {
  if (!/^#[0-9a-f]{3,8}$/i.test(c.trim())) throw new Error(`a color "#RRGGBB", not ${JSON.stringify(c)}`);
  return { solid: { color: lit(`'${c.trim()}'`) } };
};

/** What a click on a button does: a page, a bookmark, back, or a web link. */
function buttonAction(a: NonNullable<VisualSpec["action"]>) {
  const type = { page: "PageNavigation", bookmark: "Bookmark", back: "Back", url: "WebUrl" }[a.type];
  if (!type) throw new Error('action.type: "page", "bookmark", "back" or "url".');
  if (a.type !== "back" && !a.target) throw new Error(`action.target: the ${a.type === "page" ? "page name" : a.type === "bookmark" ? "bookmark name" : "link"}.`);
  return { show: lit("true"), type: textLiteral(type), ...(a.type === "page" && { navigationSection: textLiteral(a.target!) }), ...(a.type === "bookmark" && { bookmark: textLiteral(a.target!) }), ...(a.type === "url" && { webUrl: textLiteral(a.target!) }) };
}

/** visual.json of a spec; `schema` = the $schema the report's other visuals use (same version). */
export function buildVisual(spec: VisualSpec, name: string, z: number, info: ModelInfo | undefined, extension: ReportMeasures, schema = VISUAL_SCHEMA) {
  const roles = ROLES[spec.type];
  if (!roles) throw new Error(`Visual type "${spec.type}" is not built by create_visual. Types: ${Object.keys(ROLES).join(", ")} (${TYPE_NOTES}).`);
  const given = Object.entries(fitRoles(spec.type, spec.fields) ?? {}).filter(([, list]) => list?.length);
  const wrong = given.map(([role]) => role).filter((r) => !roles.includes(r));
  if (wrong.length) throw new Error(`A ${spec.type} has no role ${wrong.join(", ")}. Its roles: ${roles.join(", ") || "none"}.`);
  // Every field checked before anything is written (a wrong name is refused with the closest ones).
  const queryState = Object.fromEntries(given.map(([role, list]) => [roleKey(role), { projections: list.map((f) => (SINGLE_VALUE.includes(spec.type) ? oneValue(projection(f, info, extension), info) : projection(f, info, extension))) }]));
  const visual: Record<string, unknown> = { visualType: spec.type, drillFilterOtherVisuals: true };
  if (given.length) {
    const query: Record<string, unknown> = { queryState };
    if (spec.sort) {
      const sorted = projection(spec.sort.field, info, extension);
      query.sortDefinition = { sort: [{ field: sorted.field, direction: spec.sort.descending ? "Descending" : "Ascending" }], isDefaultSort: false };
    }
    visual.query = query;
  }
  const filterConfig = spec.top ? { filters: [topNFilter(queryState, spec.top, info, extension)] } : undefined;
  const st = spec.textStyle ?? {};
  if (spec.type === "textbox" && spec.text)
    visual.objects = {
      general: [
        {
          properties: {
            paragraphs: spec.text.split("\n").map((line) => ({
              textRuns: [{ value: line, ...(Object.keys(st).some((k) => k !== "align") && { textStyle: { ...(st.fontSize && { fontSize: `${st.fontSize}pt` }), ...(st.bold && { fontWeight: "bold" }), ...(st.italic && { fontStyle: "italic" }), ...(st.color && { color: st.color }), ...(st.fontFamily && { fontFamily: st.fontFamily }) } }) }],
              ...(st.align && { horizontalTextAlignment: st.align }),
            })),
          },
        },
      ],
    };
  if (spec.type === "image") {
    if (!spec.imageUrl || !/^https:\/\//.test(spec.imageUrl)) throw new Error("An image needs image_url: a public https link to the picture (png, jpg, svg).");
    visual.objects = { image: [{ properties: { sourceType: textLiteral("imageUrl"), sourceUrl: textLiteral(spec.imageUrl), fit: textLiteral("Fit") } }], general: [{ properties: { imageUrl: textLiteral(spec.imageUrl) } }] };
  }
  if (spec.type === "shape") {
    visual.objects = { shape: [{ properties: { tileShape: textLiteral(spec.shape ?? "rectangle") } }], ...(spec.fill && { fill: [{ properties: { show: lit("true") } }, { properties: { fillColor: colorOf(spec.fill), transparency: lit("0D") }, selector: { id: "default" } }] }), outline: [{ properties: { show: lit("false") } }] };
  }
  if (spec.type === "actionButton") {
    if (spec.icon && !BUTTON_ICONS.includes(spec.icon)) throw new Error(`No icon "${spec.icon}". Icons: ${BUTTON_ICONS.join(", ")}.`);
    const state = (props: Record<string, unknown>) => [{ properties: { show: lit("true") } }, { properties: props, selector: { id: "default" } }];
    visual.objects = {
      icon: spec.icon && spec.icon !== "blank" ? state({ shapeType: textLiteral(spec.icon), ...(st.color && { lineColor: colorOf(st.color) }) }) : [{ properties: { show: lit("false") } }],
      ...(spec.text && { text: state({ text: textLiteral(spec.text), ...(st.color && { fontColor: colorOf(st.color) }), ...(st.fontSize && { fontSize: lit(`${st.fontSize}D`) }), ...(st.bold && { bold: lit("true") }), ...(st.fontFamily && { fontFamily: textLiteral(st.fontFamily) }), ...(st.align && { horizontalAlignment: textLiteral(st.align) }) }) }),
      ...(spec.fill && { fill: state({ fillColor: colorOf(spec.fill), transparency: lit("0D") }) }),
    };
    if (spec.action) visual.visualContainerObjects = { visualLink: [{ properties: buttonAction(spec.action) }] };
  }
  if (spec.title) visual.visualContainerObjects = { ...(visual.visualContainerObjects as object), title: [{ properties: { show: lit("true"), text: textLiteral(spec.title) } }] };
  // A titled card or slicer would say its field twice (and a card's label is cut in a short box).
  if (spec.title && spec.type === "card") visual.objects = { ...(visual.objects as object), categoryLabels: [{ properties: { show: lit("false") } }] };
  if (spec.title && spec.type === "slicer") visual.objects = { ...(visual.objects as object), header: [{ properties: { show: lit("false") } }] };
  const round = (n: number) => Math.round(n);
  return { $schema: schema, name, position: { x: round(spec.x), y: round(spec.y), z, width: round(spec.width), height: round(spec.height), tabOrder: z }, visual, ...(filterConfig && { filterConfig }) };
}

/** page.json of a new page. */
export const buildPage = (name: string, displayName: string, width = 1280, height = 720, schema = PAGE_SCHEMA) => ({ $schema: schema, name, displayName, displayOption: "FitToPage", width, height });

/* --------------------------------- Formats --------------------------------- */

type Catalog = { container: string[]; visuals: Record<string, Record<string, Record<string, string>>> };
const CATALOG = catalog as Catalog;
// Buttons are not in Power BI's theme schema: their objects are a shape's, with the icon.
CATALOG.visuals.actionButton ??= {
  ...CATALOG.visuals.shape,
  icon: { show: "boolean", shapeType: BUTTON_ICONS.join("|"), lineColor: "color", lineWeight: "number", lineTransparency: "number", horizontalAlignment: "left|center|right", verticalAlignment: "top|middle|bottom", padding: "integer", iconSize: "integer" },
  text: { ...CATALOG.visuals.shape.text, horizontalAlignment: "left|center|right", verticalAlignment: "top|middle|bottom" },
};
/** Objects of buttons and shapes written per state ("default"): their show stays general. */
const STATEFUL: Record<string, string[]> = { actionButton: ["icon", "text", "fill", "outline", "glow", "shadow"], shape: ["fill", "outline", "text", "glow", "shadow"] };

const where = (o: string) => (CATALOG.container.includes(o) ? "visualContainerObjects" : "objects");
/** "title", "visualContainerObjects.title" or "objects.labels" → the object's name. */
const bareObject = (o: string) => o.trim().replace(/^(visual\.)?(visualContainerObjects|objects)\./, "");
/** How a property's value is written in format_visual (plain JSON). */
const plainType = (t: string) =>
  t === "color" ? 'color "#RRGGBB"' : t === "boolean" ? "true / false" : t === "number" || t === "integer" || t === "fontSize" ? "number" : t === "string" ? "text" : t.includes("|") ? `one of ${t.split("|").map((v) => `"${v}"`).join(", ")}` : t;

/**
 * The format options of a visual type (from Power BI's theme schema): its objects, or the properties of
 * some objects ("title, labels") with their type. Values are given plainly to format_visual.
 */
export function formatOptions(type: string, object?: string) {
  const objects = CATALOG.visuals[type];
  if (!objects) return `No format catalog for "${type}". Types: ${Object.keys(CATALOG.visuals).join(", ")}.`;
  if (!object) {
    const list = (container: boolean) => Object.keys(objects).filter((o) => CATALOG.container.includes(o) === container);
    return [`Format objects of ${type} (call again with "object", e.g. "title, labels", for their properties):`, `- visual.visualContainerObjects (the frame): ${list(true).join(", ")}`, `- visual.objects (the content): ${list(false).join(", ") || "none"}`].join("\n");
  }
  return object
    .split(",")
    .map(bareObject)
    .filter(Boolean)
    .map((o) => {
      const props = objects[o];
      if (!props) return `${type} has no object "${o}". Objects: ${Object.keys(objects).join(", ")}.`;
      return [`${o}:`, ...Object.entries(props).map(([p, t]) => `- ${p}: ${plainType(t)}`)].join("\n");
    })
    .join("\n\n");
}

/** A plain value (16, true, "#203247", "Outer") → its PBIR expression, by the property's type. */
function expression(t: string, value: unknown, field?: (spec: string) => unknown) {
  // A value from a measure (conditional formatting: a measure returning a color, a text…).
  if (value && typeof value === "object" && typeof (value as { measure?: unknown }).measure === "string") {
    if (!field) throw new Error("a measure here needs the model");
    const expr = field((value as { measure: string }).measure);
    return t === "color" ? { solid: { color: { expr } } } : { expr };
  }
  if (value && typeof value === "object") return value; // already an expression
  if (t === "color") {
    if (typeof value !== "string" || !/^#[0-9a-f]{3,8}$/i.test(value.trim())) throw new Error(`a color "#RRGGBB", not ${JSON.stringify(value)}`);
    return { solid: { color: lit(`'${value.trim()}'`) } };
  }
  if (t === "boolean") {
    if (value !== true && value !== false && value !== "true" && value !== "false") throw new Error(`true or false, not ${JSON.stringify(value)}`);
    return lit(String(value));
  }
  if (t === "number" || t === "integer" || t === "fontSize") {
    const n = Number(value);
    if (value === "" || value === null || Number.isNaN(n)) throw new Error(`a number, not ${JSON.stringify(value)}`);
    return lit(t === "integer" ? `${Math.round(n)}L` : `${n}D`);
  }
  if (t.includes("|")) {
    const allowed = t.split("|");
    const v = allowed.find((a) => a.toLowerCase() === String(value).toLowerCase());
    if (!v) throw new Error(`one of ${allowed.map((a) => `"${a}"`).join(", ")}, not ${JSON.stringify(value)}`);
    return allowed.every((a) => /^\d+$/.test(a)) ? lit(`${v}D`) : textLiteral(v);
  }
  return typeof value === "number" ? lit(`${value}D`) : typeof value === "boolean" ? lit(String(value)) : textLiteral(String(value));
}

export type Format = Record<string, Record<string, unknown>>;

/** The format objects of a visual type and their properties (name → type), from the catalog. */
export const formatObjects = (type: string): Record<string, Record<string, string>> => CATALOG.visuals[type] ?? {};

/**
 * JSON Patch operations writing a plain format ({"title": {"fontSize": 16, "bold": true}}) into a
 * visual.json: names checked against the catalog, values written as expressions, each object's entries kept and
 * completed (the general one, a series' one, a button state's). Every mistake is listed at once.
 */
export function formatOps(json: any, format: Format, field?: (spec: string) => unknown) {
  const type = String(json?.visual?.visualType ?? "");
  const objects = CATALOG.visuals[type];
  if (!objects) throw new Error(`No format catalog for "${type}": change it with update_visual.`);
  const ops: { op: "add"; path: string; value: unknown }[] = [];
  const errors: string[] = [];
  /** Names that do not exist on this type: skipped (the rest is applied), and said in the result. */
  const ignored: string[] = [];
  const lists: Record<string, any[]> = {};
  for (const [given0, props] of Object.entries(format ?? {})) {
    // "dataPoint:[Revenue]" = the series of a measure, "dataPoint:East" = a value of the visual's series (or category) column.
    const [given, only] = given0.split(/:(.+)/) as [string, string | undefined];
    const o = bareObject(given);
    const known = objects[o];
    if (!known) {
      ignored.push(`no object "${given}" on a ${type} (its objects: ${Object.keys(objects).join(", ")})`);
      continue;
    }
    if (!props || typeof props !== "object") {
      errors.push(`${o} must be an object of properties`);
      continue;
    }
    // One working copy per object: several selectors of the same object ("dataPoint:A", "dataPoint:B") add up.
    const key = `${where(o)}/${o}`;
    const list: any[] = (lists[key] ??= JSON.parse(JSON.stringify(json.visual?.[where(o)]?.[o] ?? [])));
    let selector: unknown;
    try {
      selector = only ? seriesSelector(json, only, field) : undefined;
    } catch (e) {
      errors.push(`${given0}: ${e instanceof Error ? e.message : e}`);
      continue;
    }
    /** The entry of the object for a selector (none: the general one), added at the end when missing. */
    const entry = (sel: unknown) => {
      const at = list.findIndex((e) => JSON.stringify(e?.selector ?? null) === JSON.stringify(sel ?? null));
      if (at >= 0) return list[at];
      list.push({ properties: {}, ...(sel ? { selector: sel } : {}) });
      return list.at(-1);
    };
    const stateful = !only && STATEFUL[type]?.includes(o);
    for (const [p, value] of Object.entries(props)) {
      const t = known[p];
      if (!t) {
        ignored.push(`${o}.${p} (no such property on a ${type}; ${o} has: ${Object.keys(known).join(", ")})`);
        continue;
      }
      try {
        const e = entry(stateful && p !== "show" ? { id: "default" } : selector);
        e.properties = { ...(e.properties ?? {}), [p]: expression(t, value, field) };
      } catch (e) {
        errors.push(`${o}.${p} must be ${e instanceof Error ? e.message : e}`);
      }
    }
  }
  if (errors.length) throw new Error(`${type}: ${errors.join("; ")}. Nothing was changed.`);
  for (const [key, value] of Object.entries(lists)) if (value.length) ops.push({ op: "add", path: `/visual/${key}`, value });
  return { ops, ignored };
}

/** The selector of one series: a measure of the visual ([Revenue]), or a value of its series / category column. */
function seriesSelector(json: any, only: string, field?: (spec: string) => unknown) {
  const state = json?.visual?.query?.queryState ?? {};
  const all = Object.values<any>(state).flatMap((r) => r.projections ?? []);
  const name = only.trim().match(/^\[(.+)\]$/)?.[1] ?? only.trim().match(/^'?[^'[]+'?\[(.+)\]$/)?.[1];
  if (name) {
    const hit = all.find((p) => p.field?.Measure?.Property?.toLowerCase() === name.toLowerCase() || String(p.queryRef).toLowerCase().endsWith(`.${name.toLowerCase()}`));
    if (!hit) throw new Error(`no measure [${name}] in this visual (its fields: ${all.map((p) => p.queryRef).join(", ")})`);
    return { metadata: hit.queryRef };
  }
  const column = ["Series", "Category", "Rows", "Columns"].flatMap((r) => state[r]?.projections ?? []).find((p: any) => p.field?.Column)?.field.Column;
  if (!column) throw new Error("this visual has no series or category column to pick a value of; name a measure as [Measure]");
  const value = (/^-?\d+(\.\d+)?$/.test(only.trim()) ? lit(`${only.trim()}L`) : textLiteral(only.trim())).expr;
  void field;
  return { data: [{ scopeId: { Comparison: { ComparisonKind: 0, Left: { Column: column }, Right: value } } }] };
}

/** What format results say about skipped names. */
export const ignoredText = (ignored: string[]) => (ignored.length ? ` Ignored: ${ignored.join("; ")}.` : "");

/** page.json operations for the page's colors: background (behind the visuals) and wallpaper (around the page). */
export function pageFormatOps(json: any, format: { background?: string; wallpaper?: string }) {
  const color = (c: string) => {
    if (!/^#[0-9a-f]{6}$/i.test(c)) throw new Error(`a color "#RRGGBB", not ${JSON.stringify(c)}`);
    return [{ properties: { color: { solid: { color: lit(`'${c}'`) } }, transparency: lit("0D") } }];
  };
  const objects = { ...(json.objects ?? {}) };
  if (format.background) objects.background = color(format.background);
  if (format.wallpaper) objects.outspace = color(format.wallpaper);
  return [{ op: "add" as const, path: "/objects", value: objects }];
}
