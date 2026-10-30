/**
 * Parts of a visual's PBIR the AI asks for by a short spec, written by this code: buttons and their
 * action (page navigation, back, web link), conditional formatting (gradient, rules, a color measure)
 * and reference lines (a constant or a measure). Fields are checked against the model first.
 */
import type { ModelInfo } from "@/services/model";
import { projection, textLiteral, type ReportMeasures } from "./visual-spec";

const lit = (value: string) => ({ expr: { Literal: { Value: value } } });
const color = (c: string) => {
  if (!/^#[0-9a-f]{6}$/i.test(c?.trim() ?? "")) throw new Error(`a color "#RRGGBB", not ${JSON.stringify(c)}`);
  return { Literal: { Value: `'${c.trim()}'` } };
};
const number = (n: number) => ({ Literal: { Value: `${n}D` } });

/* --------------------------------- Actions --------------------------------- */

export type Action = { type: "page" | "back" | "url"; page?: string; url?: string; tooltip?: string };

/** The visualLink of an action (any visual can carry one: a button, a shape, a card); `pageName` = the target page's name (id). */
export function linkProperties(action: Action, pageName?: string) {
  const props: Record<string, unknown> = { show: lit("true") };
  if (action.type === "page") {
    if (!pageName) throw new Error('A page action needs "page".');
    Object.assign(props, { type: textLiteral("PageNavigation"), navigationSection: textLiteral(pageName) });
  } else if (action.type === "back") props.type = textLiteral("Back");
  else {
    if (!action.url || !/^https?:\/\//i.test(action.url)) throw new Error('A url action needs "url" (https://…).');
    Object.assign(props, { type: textLiteral("WebUrl"), webUrl: textLiteral(action.url) });
  }
  if (action.tooltip) props.tooltip = textLiteral(action.tooltip);
  return [{ properties: props }];
}

/** visual.json of a button: its text, colors and action. */
export function buttonVisual(spec: { name: string; z: number; x: number; y: number; width: number; height: number; text: string; fill?: string; fontColor?: string; action: Action; pageName?: string; schema: string }) {
  const id = { selector: { id: "default" } };
  return {
    $schema: spec.schema,
    name: spec.name,
    position: { x: Math.round(spec.x), y: Math.round(spec.y), z: spec.z, width: Math.round(spec.width), height: Math.round(spec.height), tabOrder: spec.z },
    visual: {
      visualType: "actionButton",
      objects: {
        icon: [{ properties: { shapeType: textLiteral("blank") }, ...id }],
        text: [{ properties: { show: lit("true"), text: textLiteral(spec.text), ...(spec.fontColor && { fontColor: { solid: { color: { expr: color(spec.fontColor) } } } }) }, ...id }],
        ...(spec.fill && { fill: [{ properties: { show: lit("true"), fillColor: { solid: { color: { expr: color(spec.fill) } } }, transparency: lit("0D") }, ...id }] }),
      },
      visualContainerObjects: { visualLink: linkProperties(spec.action, spec.pageName) },
      drillFilterOtherVisuals: true,
    },
  };
}

/* --------------------------- Conditional formatting --------------------------- */

export type Conditional = {
  /** The column or measure of the visual whose cells / bars are colored (tables, matrices); charts: the value. */
  field: string;
  /** background / font (tables, matrices), color (bars, columns, lines, pie slices). */
  target?: "background" | "font" | "color";
  /** Based on: a value (default: `field` itself), or a measure returning a color "#RRGGBB" (from_field). */
  based_on?: string;
  gradient?: { min: string; max: string; mid?: string };
  rules?: { op: ">" | ">=" | "<" | "<=" | "="; value: number; color: string }[];
  from_field?: boolean;
};

const KIND: Record<string, number> = { "=": 0, ">": 1, ">=": 2, "<": 3, "<=": 4 };
const TABLES = ["tableEx", "pivotTable"];

/**
 * The patch writing a conditional format into a visual.json: tables and matrices color the cells of
 * `field` (background or font), charts their data points. Earlier rules on the same target are replaced.
 */
export function conditionalOps(json: any, c: Conditional, info: ModelInfo | undefined, extension: ReportMeasures) {
  const type = String(json?.visual?.visualType ?? "");
  const state = json?.visual?.query?.queryState ?? {};
  const projections = Object.values<any>(state).flatMap((s) => s.projections ?? []);
  const field = projection(c.field, info, extension);
  const inVisual = projections.find((p: any) => p.queryRef === field.queryRef);
  if (!inVisual) throw new Error(`${c.field} is not in this visual (its fields: ${projections.map((p: any) => p.queryRef).join(", ") || "none"}). Format a field it shows.`);
  const input = (c.based_on ? projection(c.based_on, info, extension) : field).field;
  let value: unknown;
  if (c.from_field) value = { expr: input };
  else if (c.gradient) {
    const g = c.gradient;
    const stop = (col: string) => ({ color: color(col) });
    const rule = g.mid ? { linearGradient3: { min: stop(g.min), mid: stop(g.mid), max: stop(g.max), nullColoringStrategy: { strategy: { Literal: { Value: "'asZero'" } } } } } : { linearGradient2: { min: stop(g.min), max: stop(g.max), nullColoringStrategy: { strategy: { Literal: { Value: "'asZero'" } } } } };
    value = { expr: { FillRule: { Input: input, FillRule: rule } } };
  } else if (c.rules?.length) {
    const cases = c.rules.map((r) => {
      if (!(r.op in KIND)) throw new Error(`Rule operator "${r.op}": one of ${Object.keys(KIND).join(" ")}.`);
      if (typeof r.value !== "number") throw new Error("A rule's value is a number.");
      return { Condition: { Comparison: { ComparisonKind: KIND[r.op], Left: input, Right: number(r.value) } }, Value: color(r.color) };
    });
    value = { expr: { Conditional: { Cases: cases } } };
  } else throw new Error('Give "gradient" ({min, max, mid?} colors), "rules" ([{op, value, color}]) or "from_field": true (a measure returning "#RRGGBB" in based_on).');
  const selector = { data: [{ dataViewWildcard: { matchingOption: 1 } }], ...(TABLES.includes(type) && { metadata: field.queryRef }) };
  let object: string;
  let property: string;
  if (TABLES.includes(type)) {
    if (c.target === "color") throw new Error('A table or matrix colors its cells: target "background" or "font".');
    object = "values";
    property = c.target === "font" ? "fontColor" : "backColor";
  } else if (/Chart$/.test(type) || type === "treemap") {
    object = "dataPoint";
    property = "fill";
  } else if (type === "card" || type === "multiRowCard") {
    object = type === "card" ? "labels" : "dataLabels";
    property = "color";
  } else throw new Error(`Conditional formatting is written here for tables, matrices, charts and cards, not a ${type}: use update_visual.`);
  const list: any[] = structuredClone(json.visual?.objects?.[object] ?? []);
  const same = (e: any) => JSON.stringify(e?.selector) === JSON.stringify(selector) && e?.properties?.[property];
  const kept = list.filter((e) => !same(e));
  const entry = { properties: { [property]: { solid: { color: value } } }, selector };
  return [{ op: "add" as const, path: `/visual/objects/${object}`, value: [...kept, entry] }];
}

/* ------------------------------ Reference lines ------------------------------ */

export type ReferenceLine = { value: number | string; axis?: "y" | "x"; label?: string; color?: string; style?: "solid" | "dashed" | "dotted"; show_label?: boolean };

/** Constant lines of a chart (a number, or a measure: a target, an average…): replaces the chart's lines on that axis. */
export function referenceLineOps(json: any, lines: ReferenceLine[], info: ModelInfo | undefined, extension: ReportMeasures) {
  const type = String(json?.visual?.visualType ?? "");
  if (!/Chart$/.test(type) || /pie|donut/i.test(type)) throw new Error(`Reference lines are on charts with axes (line, column, bar, area, combo), not a ${type}.`);
  const byAxis = new Map<string, unknown[]>();
  lines.forEach((l, i) => {
    const object = l.axis === "x" ? "xAxisReferenceLine" : "y1AxisReferenceLine";
    const value = typeof l.value === "number" ? { expr: number(l.value) } : /^-?\d+(\.\d+)?$/.test(String(l.value)) ? { expr: number(Number(l.value)) } : { expr: projection(String(l.value), info, extension).field };
    const props: Record<string, unknown> = {
      show: lit("true"),
      value,
      lineColor: { solid: { color: { expr: color(l.color ?? "#E74C3C") } } },
      style: textLiteral(l.style ?? "dashed"),
      width: lit("2D"),
      dataLabelShow: lit(String(l.show_label ?? !!l.label)),
      ...(l.label && { displayName: textLiteral(l.label), dataLabelText: textLiteral("ValueAndName") }),
    };
    byAxis.set(object, [...(byAxis.get(object) ?? []), { properties: props, selector: { id: String(i + 1) } }]);
  });
  return [...byAxis].map(([object, list]) => ({ op: "add" as const, path: `/visual/objects/${object}`, value: list }));
}
