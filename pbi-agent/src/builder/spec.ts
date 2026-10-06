/** Report being built: pages of visuals with their fields, filters and formatting. */
import { z } from "zod";
import { AGGREGATES } from "@/query/pivot";

type RoleKind = "dimension" | "value" | "any";
type Role = { label: string; kind: RoleKind; max?: number };

/** Main Power BI visuals and their data roles (Power BI role names). */
export const VISUALS = {
  card: { label: "Card", roles: { Values: { label: "Value", kind: "value", max: 1 } } },
  clusteredColumnChart: { label: "Column chart", roles: { Category: { label: "X axis", kind: "dimension", max: 1 }, Y: { label: "Y axis", kind: "value" }, Series: { label: "Legend", kind: "dimension", max: 1 } } },
  columnChart: { label: "Stacked column chart", roles: { Category: { label: "X axis", kind: "dimension", max: 1 }, Y: { label: "Y axis", kind: "value" }, Series: { label: "Legend", kind: "dimension", max: 1 } } },
  clusteredBarChart: { label: "Bar chart", roles: { Category: { label: "Y axis", kind: "dimension", max: 1 }, Y: { label: "X axis", kind: "value" }, Series: { label: "Legend", kind: "dimension", max: 1 } } },
  barChart: { label: "Stacked bar chart", roles: { Category: { label: "Y axis", kind: "dimension", max: 1 }, Y: { label: "X axis", kind: "value" }, Series: { label: "Legend", kind: "dimension", max: 1 } } },
  lineChart: { label: "Line chart", roles: { Category: { label: "X axis", kind: "dimension", max: 1 }, Y: { label: "Y axis", kind: "value" }, Series: { label: "Legend", kind: "dimension", max: 1 } } },
  areaChart: { label: "Area chart", roles: { Category: { label: "X axis", kind: "dimension", max: 1 }, Y: { label: "Y axis", kind: "value" }, Series: { label: "Legend", kind: "dimension", max: 1 } } },
  pieChart: { label: "Pie chart", roles: { Category: { label: "Legend", kind: "dimension", max: 1 }, Y: { label: "Values", kind: "value", max: 1 } } },
  donutChart: { label: "Donut chart", roles: { Category: { label: "Legend", kind: "dimension", max: 1 }, Y: { label: "Values", kind: "value", max: 1 } } },
  scatterChart: { label: "Scatter chart", roles: { Category: { label: "Values", kind: "dimension", max: 1 }, X: { label: "X axis", kind: "value", max: 1 }, Y: { label: "Y axis", kind: "value", max: 1 } } },
  tableEx: { label: "Table", roles: { Values: { label: "Columns", kind: "any" } } },
  pivotTable: { label: "Matrix", roles: { Rows: { label: "Rows", kind: "dimension" }, Columns: { label: "Columns", kind: "dimension" }, Values: { label: "Values", kind: "value" } } },
  slicer: { label: "Slicer", roles: { Values: { label: "Field", kind: "dimension", max: 1 } } },
  gauge: { label: "Gauge", roles: { Y: { label: "Value", kind: "value", max: 1 }, MaxValue: { label: "Maximum", kind: "value", max: 1 }, TargetValue: { label: "Target", kind: "value", max: 1 } } },
} satisfies Record<string, { label: string; roles: Record<string, Role> }>;

export type VisualType = keyof typeof VISUALS;
export const VISUAL_TYPES = Object.keys(VISUALS) as VisualType[];
export const rolesOf = (type: VisualType) => VISUALS[type].roles as Record<string, Role>;

const field = z.object({
  /** 'Table'[Column] or [Measure]. */
  ref: z.string(),
  /** For a column in a value role. */
  aggregate: z.enum(AGGREGATES).optional(),
});
export type Field = z.infer<typeof field>;

export const formatSchema = z.object({
  showTitle: z.boolean().default(true),
  titleSize: z.number().optional(),
  /** Data colors, in series order (hex). */
  colors: z.array(z.string()).default([]),
  showLegend: z.boolean().default(true),
  legendPosition: z.enum(["Top", "Bottom", "Left", "Right"]).default("Top"),
  showDataLabels: z.boolean().default(false),
  showAxisTitles: z.boolean().default(true),
  fontSize: z.number().optional(),
  background: z.string().optional(),
  /** Sort by a role field: "Y desc", "Category asc"… */
  sort: z.string().optional(),
});
export type Format = z.infer<typeof formatSchema>;

export const visualSchema = z.object({
  id: z.string(),
  type: z.enum(VISUAL_TYPES as [VisualType, ...VisualType[]]),
  title: z.string().default(""),
  fields: z.record(z.string(), z.array(field)).default({}),
  filters: z.array(z.object({ ref: z.string(), values: z.array(z.union([z.string(), z.number(), z.boolean()])) })).default([]),
  format: formatSchema.default(formatSchema.parse({})),
  /** Position on the page (1280 × 720). */
  x: z.number().default(0),
  y: z.number().default(0),
  width: z.number().default(600),
  height: z.number().default(320),
});
export type BuilderVisual = z.infer<typeof visualSchema>;

export const reportSchema = z.object({
  id: z.string(),
  title: z.string(),
  datasetId: z.string(),
  /** Reports of the database given to the agent as examples. */
  contextReports: z.array(z.string()).default([]),
  pages: z.array(z.object({ name: z.string(), displayName: z.string(), visuals: z.array(visualSchema) })),
  /** Copy imported in My workspace for the real Power BI preview. */
  published: z.object({ reportId: z.string(), at: z.string() }).optional(),
});
export type BuilderReport = z.infer<typeof reportSchema>;

export const hexId = () => crypto.randomUUID().replace(/-/g, "").slice(0, 20);

export const PALETTE = ["#118DFF", "#12239E", "#E66C37", "#6B007B", "#E044A7", "#744EC2", "#D9B300", "#D64550"];

export function newVisual(type: VisualType, index: number): BuilderVisual {
  return visualSchema.parse({ id: hexId(), type, title: VISUALS[type].label, ...slot(index) });
}

/** Default position: two visuals per row on a 1280 × 720 page. */
export const slot = (index: number) => ({ x: 20 + (index % 2) * 630, y: 20 + Math.floor(index / 2) * 340, width: 610, height: 320 });

export const newReport = (title: string, datasetId: string): BuilderReport => ({
  id: crypto.randomUUID(),
  title,
  datasetId,
  contextReports: [],
  pages: [{ name: `ReportSection${hexId()}`, displayName: "Page 1", visuals: [] }],
});
