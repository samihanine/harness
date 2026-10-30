/**
 * The look of a visual from a short spec the AI fills (title, field names, colors, text size, frame),
 * written by code: format objects where this visual type has them (the catalog), field display names in
 * its query, a color per series. The AI never writes format expressions itself.
 */
import { z } from "zod";
import { formatObjects, formatOps, type Format } from "./visual-spec";

const hexColor = z.string().regex(/^#[0-9a-f]{6}$/i, 'a color "#RRGGBB"');

export const styleSpec = z.object({
  title: z.string().optional().describe("Title text (shown)"),
  title_size: z.number().optional().describe("Title font size (pt)"),
  title_color: hexColor.optional(),
  title_bold: z.boolean().optional(),
  title_align: z.enum(["left", "center", "right"]).optional(),
  rename: z.record(z.string(), z.string()).optional().describe('New display names of the visual\'s fields: {"[Revenue]": "Sales", "\'Date\'[Month]": "Month"}'),
  colors: z.record(z.string(), hexColor).optional().describe('Series colors: {"[Revenue]": "#1F6FEB", "[Cost]": "#D29922"}; "*" = every data point (one-series charts)'),
  text_size: z.number().optional().describe("Font size (pt) of the content: data labels, legend, axes, table values and headers, card labels"),
  text_color: hexColor.optional().describe("Color of that content text"),
  labels: z.boolean().optional().describe("Data labels on / off (charts)"),
  legend: z.boolean().optional().describe("Legend on / off"),
  header_color: hexColor.optional().describe("Tables / matrices: background of the column headers"),
  header_text_color: hexColor.optional(),
  background: hexColor.optional().describe("Frame background"),
  border: hexColor.optional().describe("Frame border color (shown)"),
  radius: z.number().optional().describe("Frame corner radius (px)"),
});
export type StyleSpec = z.infer<typeof styleSpec>;

/** Content objects that hold text, with their size and color property names. */
const TEXT: [object: string, color: string][] = [["labels", "color"], ["legend", "labelColor"], ["categoryAxis", "labelColor"], ["valueAxis", "labelColor"], ["values", "fontColor"], ["columnHeaders", "fontColor"], ["rowHeaders", "fontColor"], ["categoryLabels", "color"], ["items", "fontColor"]];

/** "'Sales'[Revenue]", "Sales[Revenue]", "[Revenue]", "Sum(Sales.Amount)" → comparable keys. */
const keys = (given: string) => {
  const g = given.trim().toLowerCase();
  const m = g.match(/^'?([^'[]*)'?\[([^\]]+)\]$/);
  return m ? [m[1] ? `${m[1]}.${m[2]}` : m[2], m[2]] : [g, g.replace(/^.*\./, "").replace(/\)$/, "")];
};

/** The plain format (format_visual's form) of a spec, for the objects this visual type has. */
export function styleFormat(type: string, spec: StyleSpec): Format {
  const objects = formatObjects(type);
  const format: Format = {};
  const set = (o: string, p: string, v: unknown) => {
    if (v === undefined || !objects[o]?.[p]) return;
    (format[o] ??= {})[p] = v;
  };
  if (spec.title !== undefined) (set("title", "show", true), set("title", "text", spec.title));
  set("title", "fontSize", spec.title_size);
  set("title", "fontColor", spec.title_color);
  set("title", "bold", spec.title_bold);
  set("title", "alignment", spec.title_align);
  for (const [o, color] of TEXT) {
    set(o, "fontSize", spec.text_size);
    set(o, color, spec.text_color);
  }
  set("labels", "show", spec.labels);
  set("legend", "show", spec.legend);
  set("columnHeaders", "backColor", spec.header_color);
  set("columnHeaders", "fontColor", spec.header_text_color);
  if (spec.background) (set("background", "show", true), set("background", "color", spec.background));
  if (spec.border) (set("border", "show", true), set("border", "color", spec.border));
  set("border", "radius", spec.radius);
  if (spec.colors?.["*"]) set("dataPoint", "defaultColor", spec.colors["*"]);
  return format;
}

/** JSON Patch operations of a spec on a visual.json (PBIR): its format, field names and series colors. */
export function styleOps(json: any, spec: StyleSpec) {
  const type = String(json?.visual?.visualType ?? "");
  const { ops, ignored } = formatOps(json, styleFormat(type, spec));
  const notes = [...ignored];
  const state = json.visual?.query?.queryState ?? {};
  const projections = Object.entries(state).flatMap(([role, r]: [string, any]) => (r?.projections ?? []).map((p: any, i: number) => ({ role, i, p })));
  const find = (given: string) => {
    const want = keys(given);
    return projections.find(({ p }) => [p.queryRef, p.nativeQueryRef, p.displayName].some((k) => k && want.includes(String(k).toLowerCase())));
  };
  for (const [field, name] of Object.entries(spec.rename ?? {})) {
    const found = find(field);
    if (!found) notes.push(`no field "${field}" in this visual (its fields: ${projections.map(({ p }) => p.queryRef).join(", ")})`);
    else ops.push({ op: "add", path: `/visual/query/queryState/${found.role}/projections/${found.i}/displayName`, value: name });
  }
  const series = Object.entries(spec.colors ?? {}).filter(([k]) => k !== "*");
  if (series.length) {
    if (!formatObjects(type).dataPoint) notes.push(`a ${type} has no series colors`);
    else {
      // One dataPoint entry per series, selected by its query reference; the others are kept.
      const entries = [...(json.visual?.objects?.dataPoint ?? [])];
      for (const [field, color] of series) {
        const found = find(field);
        if (!found) {
          notes.push(`no series "${field}" in this visual`);
          continue;
        }
        const entry = { properties: { fill: { solid: { color: { expr: { Literal: { Value: `'${color}'` } } } } } }, selector: { metadata: found.p.queryRef } };
        const at = entries.findIndex((e) => e?.selector?.metadata === found.p.queryRef);
        if (at >= 0) entries[at] = entry;
        else entries.push(entry);
      }
      ops.push({ op: "add", path: "/visual/objects/dataPoint", value: entries });
    }
  }
  return { ops, notes };
}
