/** Power BI report definition (Report/Layout of a .pbix) from a builder report. */
import type { ModelContent } from "@/pbi/types";
import { parseField } from "@/query/pivot";
import type { BuilderReport, BuilderVisual, Field } from "./spec";
import { PALETTE, rolesOf } from "./spec";

const AGG = { sum: 0, avg: 1, count: 2, min: 3, max: 4, distinctcount: 5 } as const;
const AGG_NAME = { sum: "Sum", avg: "Avg", count: "Count", min: "Min", max: "Max", distinctcount: "CountNonNull" } as const;

/** Literal expression: 'text', true, 12D. */
const lit = (value: string | number | boolean) => ({
  expr: { Literal: { Value: typeof value === "string" ? `'${value.replace(/'/g, "''")}'` : typeof value === "number" ? `${value}D` : String(value) } },
});
const color = (hex: string) => ({ solid: { color: lit(hex) } });

/** Table of a measure (layouts reference measures through their home table). */
const measureTable = (model: ModelContent, name: string) => model.tables.find((t) => t.measures.some((m) => m.name === name))?.name ?? model.tables[0]?.name ?? "";

function visualConfig(visual: BuilderVisual, model: ModelContent) {
  const sources = new Map<string, string>();
  const source = (entity: string) => {
    if (!sources.has(entity)) sources.set(entity, `s${sources.size}`);
    return sources.get(entity)!;
  };
  const select: Record<string, unknown>[] = [];
  const projections: Record<string, { queryRef: string }[]> = {};
  const roles = rolesOf(visual.type);

  const item = (f: Field, role: string) => {
    const { table, column } = parseField(f.ref);
    if (!table) {
      const home = measureTable(model, column);
      return { Measure: { Expression: { SourceRef: { Source: source(home) } }, Property: column }, Name: `${home}.${column}` };
    }
    const col = { Column: { Expression: { SourceRef: { Source: source(table) } }, Property: column } };
    const aggregate = f.aggregate ?? (roles[role]?.kind === "value" ? "sum" : undefined);
    if (!aggregate) return { ...col, Name: `${table}.${column}` };
    return { Aggregation: { Expression: col, Function: AGG[aggregate] }, Name: `${AGG_NAME[aggregate]}(${table}.${column})` };
  };

  for (const [role, fields] of Object.entries(visual.fields)) {
    projections[role] = [];
    for (const f of fields) {
      const entry = item(f, role);
      if (!select.some((s) => s.Name === entry.Name)) select.push(entry);
      projections[role].push({ queryRef: entry.Name });
    }
  }

  const fmt = visual.format;
  const objects: Record<string, unknown[]> = {
    legend: [{ properties: { show: lit(fmt.showLegend), position: lit(fmt.legendPosition) } }],
    labels: [{ properties: { show: lit(fmt.showDataLabels), ...(fmt.fontSize ? { fontSize: lit(fmt.fontSize) } : {}) } }],
    categoryAxis: [{ properties: { showAxisTitle: lit(fmt.showAxisTitles) } }],
    valueAxis: [{ properties: { showAxisTitle: lit(fmt.showAxisTitles) } }],
  };
  if (fmt.colors.length) objects.dataPoint = [{ properties: { defaultColor: color(fmt.colors[0] ?? PALETTE[0]), fill: color(fmt.colors[0]) } }];
  const vcObjects: Record<string, unknown[]> = {
    title: [{ properties: { show: lit(fmt.showTitle), text: lit(visual.title), ...(fmt.titleSize ? { fontSize: lit(fmt.titleSize) } : {}) } }],
  };
  if (fmt.background) vcObjects.background = [{ properties: { show: lit(true), color: color(fmt.background) } }];

  const position = { x: visual.x, y: visual.y, z: 0, width: visual.width, height: visual.height };
  return {
    name: visual.id,
    layouts: [{ id: 0, position }],
    singleVisual: {
      visualType: visual.type,
      projections,
      prototypeQuery: { Version: 2, From: [...sources].map(([Entity, Name]) => ({ Name, Entity, Type: 0 })), Select: select },
      drillFilterOtherVisuals: true,
      objects,
      vcObjects,
    },
  };
}

/** Visual-level basic filters ("is one of"). */
function visualFilters(visual: BuilderVisual) {
  return visual.filters
    .filter((f) => f.values.length)
    .map((f, i) => {
      const { table, column } = parseField(f.ref);
      const literal = (v: string | number | boolean) => lit(v).expr;
      return {
        name: `Filter${i}${visual.id.slice(0, 8)}`,
        expression: { Column: { Expression: { SourceRef: { Entity: table } }, Property: column } },
        filter: {
          Version: 2,
          From: [{ Name: "f", Entity: table, Type: 0 }],
          Where: [
            {
              Condition: {
                In: { Expressions: [{ Column: { Expression: { SourceRef: { Source: "f" } }, Property: column } }], Values: f.values.map((v) => [literal(v)]) },
              },
            },
          ],
        },
        type: "Categorical",
        howCreated: 1,
      };
    });
}

export function buildLayout(report: BuilderReport, model: ModelContent) {
  return {
    id: 0,
    sections: report.pages.map((page, ordinal) => ({
      name: page.name,
      displayName: page.displayName,
      ordinal,
      width: 1280,
      height: 720,
      displayOption: 1,
      config: "{}",
      filters: "[]",
      visualContainers: page.visuals.map((visual, z) => ({
        x: visual.x,
        y: visual.y,
        z: z * 1000,
        width: visual.width,
        height: visual.height,
        config: JSON.stringify(visualConfig(visual, model)),
        filters: JSON.stringify(visualFilters(visual)),
      })),
    })),
    config: JSON.stringify({ version: "5.43", activeSectionIndex: 0, defaultDrillFilterOtherVisuals: true, settings: { useStylableVisualContainerHeader: true } }),
    layoutOptimization: 0,
    filters: "[]",
  };
}
