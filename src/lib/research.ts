/**
 * A research: the semantic models it reads, a canvas (markdown and Power BI visuals) and, per model,
 * one Power BI report whose pages each hold a short text and ONE visual (a matrix by default, or a
 * chart) — the canvas shows these visuals. The AI fills short specs; this code writes the report
 * definition (PBIR), with the research's own report measures, and creates the report.
 */
import { all } from "./store";
import type { Ref, Rows } from "./pbi";
import { fabric } from "./ms";
import { resolveField, type ModelInfo } from "./model";
import { definition, textPayload, writeDefinition, type Part } from "./report-def";
import { workspaceId } from "./tmdl";
import { buildPage, buildVisual, newName, type ReportMeasures } from "./visual-spec";

/** The visual of a research page: a matrix (default) or a chart. */
export const CHARTS = ["matrix", "table", "bar", "column", "line", "area", "pie", "card"] as const;
export type ChartKind = (typeof CHARTS)[number];

/**
 * A block of the canvas: text, or the visual of a page of a model's research report (`page`: its
 * title). "chart" blocks are older researches' DAX charts (shown as such, no longer made).
 */
export type Block =
  | { id: string; kind: "markdown"; text: string }
  | { id: string; kind: "visual"; datasetId: string; page: string }
  | { id: string; kind: "chart"; title: string; chart: string; datasetId: string; dax: string; rows?: Rows; error?: string; at?: number };

/**
 * A page of a research report: a short text (optional) above ONE visual — rows (matrix rows / chart
 * axis), columns (matrix columns / chart series), values — and page filters. `name` / `visualName`:
 * kept from one save to the next (the canvas embeds that visual).
 */
export type MatrixPage = { title: string; text?: string; chart?: ChartKind; rows: string[]; columns: string[]; values: string[]; filters: { field: string; values: (string | number)[] }[]; name?: string; visualName?: string };

/** A measure of the research report (not of the model): ratios, comparisons… usable in its pages and queries. */
export type ResearchMeasure = { table: string; name: string; expression: string; format?: string };

/** `id`: the report once created; `dirty`: pages changed since it was last written. */
export type ResearchReport = { id?: string; groupId?: string; pages: MatrixPage[]; measures?: ResearchMeasure[]; dirty?: boolean; saved?: number };

export type ResearchEntry = { id: string; name: string; datasetIds: string[]; blocks: Block[]; reports: Record<string, ResearchReport | undefined>; updatedAt: number };

export const newResearch = (name: string): ResearchEntry => ({ id: crypto.randomUUID(), name, datasetIds: [], blocks: [], reports: {}, updatedAt: Date.now() });

/** Older researches (Excel pivots) become empty canvases of the same models. */
export const readResearch = (r: Partial<ResearchEntry> & { id: string; name: string }): ResearchEntry => ({ datasetIds: [], updatedAt: Date.now(), ...r, blocks: Array.isArray(r.blocks) ? r.blocks : [], reports: r.reports && typeof r.reports === "object" ? r.reports : {} });

/* ----------------------------------- Report ----------------------------------- */

const REPORT_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json";
const VERSION_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json";
const PAGES_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json";

const json = (path: string, value: unknown): Part => ({ path, payload: textPayload(JSON.stringify(value, null, 2)), payloadType: "InlineBase64" });

/** The literal of a filter value as Power BI writes it. */
const literal = (v: string | number) => (typeof v === "number" ? (Number.isInteger(v) ? `${v}L` : `${v}D`) : `'${v.replace(/'/g, "''")}'`);

/** A page filter keeping some values of a column. */
function pageFilter(f: MatrixPage["filters"][number], info: ModelInfo | undefined) {
  const known = resolveField(info, f.field);
  if (known.kind === "measure") throw new Error(`Filters are on columns: ${f.field} is a measure.`);
  const type = info?.tables.find((t) => t.name === known.table)?.columns.find((c) => c.name === known.name)?.type?.toLowerCase() ?? "string";
  const numeric = ["int64", "double", "decimal"].includes(type);
  const values = f.values.map((v) => (numeric && typeof v === "string" && /^-?\d+(\.\d+)?$/.test(v) ? Number(v) : v));
  return {
    name: newName(),
    field: { Column: { Expression: { SourceRef: { Entity: known.table } }, Property: known.name } },
    type: "Categorical",
    filter: {
      Version: 2,
      From: [{ Name: "t", Entity: known.table, Type: 0 }],
      Where: [{ Condition: { In: { Expressions: [{ Column: { Expression: { SourceRef: { Source: "t" } }, Property: known.name } }], Values: values.map((v) => [{ Literal: { Value: literal(v) } }]) } } }],
    },
  };
}

/** The visual type and roles of a page's visual. */
function visualFields(p: MatrixPage): { type: string; fields: Record<string, string[]> } {
  switch (p.chart ?? "matrix") {
    case "table":
      return { type: "tableEx", fields: { Values: [...p.rows, ...p.columns, ...p.values] } };
    case "card":
      return p.values.length > 1 ? { type: "multiRowCard", fields: { Values: p.values } } : { type: "card", fields: { Values: p.values } };
    case "pie":
      return { type: "donutChart", fields: { Category: p.rows, Y: p.values } };
    case "bar":
    case "column":
    case "line":
    case "area": {
      const type = ({ bar: "clusteredBarChart", column: "clusteredColumnChart", line: "lineChart", area: "areaChart" } as Record<string, string>)[p.chart!];
      return { type, fields: { Category: p.rows, Y: p.values, Series: p.columns } };
    }
    default:
      return { type: "pivotTable", fields: { Rows: p.rows, Columns: p.columns, Values: p.values } };
  }
}

const own = (measures: ResearchMeasure[] = []): ReportMeasures => measures.map((m) => ({ table: m.table, name: m.name }));

/** Checks a page against the model and the research's measures (wrong fields are refused with the closest names). */
export function checkPage(page: MatrixPage, info: ModelInfo | undefined, measures?: ResearchMeasure[]) {
  if (!page.values.length && !page.rows.length) throw new Error("A page's visual needs rows or values.");
  const { type, fields } = visualFields(page);
  if (type !== "pivotTable" && type !== "tableEx" && !page.values.length) throw new Error(`A ${page.chart} needs values.`);
  if (/Chart$/.test(type) && !page.rows.length) throw new Error(`A ${page.chart} chart needs rows (its axis / slices).`);
  buildVisual({ type, x: 0, y: 0, width: 1280, height: 720, fields: Object.fromEntries(Object.entries(fields).filter(([, l]) => l.length)) }, "check", 0, info, own(measures));
  page.filters.forEach((f) => pageFilter(f, info));
}

/** A markdown text as a textbox's paragraphs (**bold** kept, the other marks dropped). */
function paragraphs(text: string) {
  return text
    .split("\n")
    .map((line) => line.replace(/^\s*(#+|[-*])\s+/, (m) => (/[-*]/.test(m) ? "• " : "")).trimEnd())
    .filter((line, i, all) => line || (i > 0 && all[i - 1]))
    .map((line) => ({ textRuns: line.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((part) => (part.startsWith("**") ? { value: part.slice(2, -2), textStyle: { fontWeight: "bold" } } : { value: part.replace(/[*_`]/g, "") })) }));
}

/** The text height above the visual: a few lines. */
const textHeight = (text?: string) => (text?.trim() ? Math.min(200, 40 + 22 * text.trim().split("\n").length) : 0);

/** The pages, visuals and report measures parts of a research report (page and visual names kept). */
export function pageParts(pages: MatrixPage[], info: ModelInfo | undefined, measures: ResearchMeasure[] = []): Part[] {
  const named = pages.map((p) => ({ ...p, name: p.name ?? newName(), visualName: p.visualName ?? newName() }));
  return [
    json("definition/pages/pages.json", { $schema: PAGES_SCHEMA, pageOrder: named.map((p) => p.name), activePageName: named[0]?.name }),
    ...named.flatMap((p) => {
      const page = { ...buildPage(p.name, p.title), ...(p.filters.length && { filterConfig: { filters: p.filters.map((f) => pageFilter(f, info)) } }) };
      const { type, fields } = visualFields(p);
      const top = textHeight(p.text);
      const visual = buildVisual({ type, x: 20, y: top + 20, width: 1240, height: 680 - top, title: p.title, fields: Object.fromEntries(Object.entries(fields).filter(([, l]) => l.length)) }, p.visualName, 1000, info, own(measures));
      const parts = [json(`definition/pages/${p.name}/page.json`, page), json(`definition/pages/${p.name}/visuals/${visual.name}/visual.json`, visual)];
      if (top) {
        const box = buildVisual({ type: "textbox", x: 20, y: 10, width: 1240, height: top }, `${p.visualName.slice(0, 16)}text`, 2000, info, []);
        (box.visual as any).objects = { general: [{ properties: { paragraphs: paragraphs(p.text!) } }] };
        parts.push(json(`definition/pages/${p.name}/visuals/${box.name}/visual.json`, box));
      }
      return parts;
    }),
    ...(measures.length
      ? [
          json("definition/reportExtensions.json", {
            $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/reportExtension/1.0.0/schema.json",
            name: "extension",
            entities: [...new Set(measures.map((m) => m.table))].map((table) => ({ name: table, measures: measures.filter((m) => m.table === table).map((m) => ({ name: m.name, dataType: "Double", expression: m.expression, ...(m.format && { formatString: m.format }) })) })),
          }),
        ]
      : []),
  ];
}

/**
 * The connection to the model: copied from a report of the app on the same model when there is one
 * (exactly what Power BI wrote), else the short form Fabric accepts.
 */
async function connection(dataset: Ref): Promise<Part> {
  const report = (await all("reports")).find((r) => r.datasetId === dataset.id && r.access === "edit");
  if (report) {
    const part = await definition(report)
      .then((parts) => parts.find((p) => p.path === "definition.pbir"))
      .catch(() => undefined);
    if (part) return part;
  }
  // The $schema is required: without it Fabric reads the old format (and asks for five more properties).
  return json("definition.pbir", { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json", version: "4.0", datasetReference: { byConnection: { connectionString: `semanticmodelid=${dataset.id}` } } });
}

/**
 * A new report in PBIR format (editable by the Builder's tools), live-connected to a semantic model,
 * in a workspace (default: My workspace); returns its id.
 */
async function createReport(name: string, dataset: Ref, pagesParts: Part[], groupId?: string) {
  const ws = await workspaceId(groupId);
  const parts = [
    await connection(dataset),
    json("definition/version.json", { $schema: VERSION_SCHEMA, version: "2.0.0" }),
    json("definition/report.json", { $schema: REPORT_SCHEMA, themeCollection: { baseTheme: { name: "CY24SU10", reportVersionAtImport: "5.61", type: "SharedResources" } }, layoutOptimization: "None", settings: { useStylableVisualContainerHeader: true, defaultDrillFilterOtherVisuals: true } }),
    ...pagesParts,
  ];
  const created = await fabric(`/workspaces/${ws}/reports`, "POST", { displayName: name, definition: { parts } });
  const id = (created?.id as string | undefined) ?? (await fabric(`/workspaces/${ws}/reports`)).value.find((r: { displayName: string }) => r.displayName === name)?.id;
  if (!id) throw new Error("The report was created but could not be found: look for it in the workspace.");
  return id as string;
}

/** An empty report (one blank page) on a model, in My workspace: the Builder's AI fills it. */
export function createBlankReport(name: string, dataset: Ref) {
  const page = newName();
  return createReport(name, dataset, [
    json("definition/pages/pages.json", { $schema: PAGES_SCHEMA, pageOrder: [page], activePageName: page }),
    json(`definition/pages/${page}/page.json`, buildPage(page, "Page 1")),
  ]);
}

/** Deletes a research's reports in Power BI (when the research is reset or deleted). */
export async function deleteResearchReports(r: ResearchEntry) {
  for (const rep of Object.values(r.reports)) if (rep?.id) await fabric(`/workspaces/${await workspaceId(rep.groupId)}/reports/${rep.id}`, "DELETE").catch(() => undefined);
}

/** Creates the report of a model (in the model's workspace) or rewrites its pages. */
export async function saveResearchReport(name: string, dataset: Ref & { info?: ModelInfo }, report: ResearchReport | undefined, pages: MatrixPage[]): Promise<ResearchReport> {
  // Names given once and kept: the canvas embeds each page's visual by them.
  const named = pages.map((p) => ({ ...p, name: p.name ?? newName(), visualName: p.visualName ?? newName() }));
  const pagesParts = pageParts(named, dataset.info, report?.measures);
  if (report?.id) {
    // Kept: the connection, the report settings and the theme; the pages and measures are replaced.
    const ref = { id: report.id, groupId: report.groupId };
    const old = await definition(ref);
    await writeDefinition(ref, [...old.filter((p) => !p.path.startsWith("definition/pages/") && p.path !== "definition/reportExtensions.json"), ...pagesParts]);
    return { ...report, pages: named, dirty: false, saved: Date.now() };
  }
  const id = await createReport(name, dataset, pagesParts, dataset.groupId);
  // My workspace: no group (the Power BI API refuses its id as a group).
  return { ...report, id, groupId: dataset.groupId, pages: named, dirty: false, saved: Date.now() };
}
