/**
 * A research: the semantic models it reads, a canvas (markdown and charts computed by DAX) and, per
 * model, one Power BI report whose pages are each a full-page matrix (rows, columns, values, filters).
 * The AI fills short specs; this code writes the report definition (PBIR) and creates the report.
 */
import { all } from "./store";
import { dax, type Ref, type Rows } from "./pbi";
import { fabric } from "./ms";
import { resolveField, type ModelInfo } from "./model";
import { definition, textPayload, writeDefinition, type Part } from "./report-def";
import { workspaceId } from "./tmdl";
import { buildPage, buildVisual, newName } from "./visual-spec";

export const CHARTS = ["bar", "column", "line", "area", "pie", "kpi", "table"] as const;
export type ChartKind = (typeof CHARTS)[number];

/** A block of the canvas: text, or a chart of a DAX query's rows (first column = categories). */
export type Block =
  | { id: string; kind: "markdown"; text: string }
  | { id: string; kind: "chart"; title: string; chart: ChartKind; datasetId: string; dax: string; rows?: Rows; error?: string; at?: number };

/** A page of a research report: one matrix filling the page. */
export type MatrixPage = { title: string; rows: string[]; columns: string[]; values: string[]; filters: { field: string; values: (string | number)[] }[] };

/** `id`: the report once created; `dirty`: pages changed since it was last written. */
export type ResearchReport = { id?: string; groupId?: string; pages: MatrixPage[]; dirty?: boolean; saved?: number };

export type ResearchEntry = { id: string; name: string; datasetIds: string[]; blocks: Block[]; reports: Record<string, ResearchReport | undefined>; updatedAt: number };

export const newResearch = (name: string): ResearchEntry => ({ id: crypto.randomUUID(), name, datasetIds: [], blocks: [], reports: {}, updatedAt: Date.now() });

/** Older researches (Excel pivots) become empty canvases of the same models. */
export const readResearch = (r: Partial<ResearchEntry> & { id: string; name: string }): ResearchEntry => ({ datasetIds: [], updatedAt: Date.now(), ...r, blocks: Array.isArray(r.blocks) ? r.blocks : [], reports: r.reports && typeof r.reports === "object" ? r.reports : {} });

/** Runs a chart's query (a chart without rows shows nothing). */
export async function computeBlock(block: Extract<Block, { kind: "chart" }>, dataset: Ref): Promise<Extract<Block, { kind: "chart" }>> {
  try {
    const rows = await dax(dataset, block.dax);
    // Column names as "Table[Name]" or "[Name]" → "Name".
    const clean = rows.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k.replace(/^.*\[(.+)\]$/, "$1"), v])));
    return { ...block, rows: clean.slice(0, 500), error: undefined, at: Date.now() };
  } catch (e) {
    return { ...block, rows: undefined, error: e instanceof Error ? e.message : String(e) };
  }
}

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

/** Checks a page against the model (wrong fields are refused with the closest names). */
export function checkPage(page: MatrixPage, info: ModelInfo | undefined) {
  if (!page.values.length && !page.rows.length) throw new Error("A matrix needs rows or values.");
  buildVisual({ type: "pivotTable", x: 0, y: 0, width: 1280, height: 720, fields: { Rows: page.rows, Columns: page.columns, Values: page.values } }, "check", 0, info, []);
  page.filters.forEach((f) => pageFilter(f, info));
}

/** The pages and visuals parts of a research report. */
function pageParts(pages: MatrixPage[], info: ModelInfo | undefined): Part[] {
  const names = pages.map(() => newName());
  return [
    json("definition/pages/pages.json", { $schema: PAGES_SCHEMA, pageOrder: names, activePageName: names[0] }),
    ...pages.flatMap((p, i) => {
      const page = { ...buildPage(names[i], p.title), ...(p.filters.length && { filterConfig: { filters: p.filters.map((f) => pageFilter(f, info)) } }) };
      const visual = buildVisual({ type: "pivotTable", x: 0, y: 0, width: 1280, height: 720, title: p.title, fields: { Rows: p.rows, Columns: p.columns, Values: p.values } }, newName(), 0, info, []);
      return [json(`definition/pages/${names[i]}/page.json`, page), json(`definition/pages/${names[i]}/visuals/${visual.name}/visual.json`, visual)];
    }),
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

/** Deletes a research's reports in Power BI (when the research is reset or deleted). */
export async function deleteResearchReports(r: ResearchEntry) {
  for (const rep of Object.values(r.reports)) if (rep?.id) await fabric(`/workspaces/${await workspaceId(rep.groupId)}/reports/${rep.id}`, "DELETE").catch(() => undefined);
}

/** Creates the report of a model (in the model's workspace) or rewrites its pages. */
export async function saveResearchReport(name: string, dataset: Ref & { info?: ModelInfo }, report: ResearchReport | undefined, pages: MatrixPage[]): Promise<ResearchReport> {
  const pagesParts = pageParts(pages, dataset.info);
  if (report?.id) {
    // Kept: the connection, the report settings and the theme; the pages are replaced.
    const ref = { id: report.id, groupId: report.groupId };
    const old = await definition(ref);
    await writeDefinition(ref, [...old.filter((p) => !p.path.startsWith("definition/pages/")), ...pagesParts]);
    return { ...report, pages, dirty: false, saved: Date.now() };
  }
  const ws = await workspaceId(dataset.groupId);
  const parts = [
    await connection(dataset),
    json("definition/version.json", { $schema: VERSION_SCHEMA, version: "2.0.0" }),
    json("definition/report.json", { $schema: REPORT_SCHEMA, themeCollection: { baseTheme: { name: "CY24SU10", reportVersionAtImport: "5.61", type: "SharedResources" } }, layoutOptimization: "None", settings: { useStylableVisualContainerHeader: true, defaultDrillFilterOtherVisuals: true } }),
    ...pagesParts,
  ];
  const created = await fabric(`/workspaces/${ws}/reports`, "POST", { displayName: name, definition: { parts } });
  const id = (created?.id as string | undefined) ?? (await fabric(`/workspaces/${ws}/reports`)).value.find((r: { displayName: string }) => r.displayName === name)?.id;
  if (!id) throw new Error("The report was created but could not be found: look for it in the workspace.");
  // My workspace: no group (the Power BI API refuses its id as a group).
  return { id, groupId: dataset.groupId, pages, dirty: false, saved: Date.now() };
}
