/** Power BI: links, DAX, model structure, embedding and what the AI sees of a report page. */
import type * as PbiClient from "powerbi-client";
import { pbi, powerBiToken } from "./ms";

export type Ref = { id: string; groupId?: string };
export type Report = PbiClient.Report;
type Visual = PbiClient.VisualDescriptor;

const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
export const scope = (groupId?: string) => (groupId ? `/groups/${groupId}` : "");

/** app.powerbi.com/groups/<ws>/reports/<id>/… (or a bare id) → ids. */
export function parseReportUrl(url: string): Ref {
  const id = url.match(/reports\/([0-9a-f-]{36})/i)?.[1] ?? url.match(/reportId=([0-9a-f-]{36})/i)?.[1] ?? url.match(GUID)?.[0];
  if (!id) throw new Error("No report id in this link");
  const group = url.match(/groups\/([0-9a-f-]{36})/i)?.[1] ?? url.match(/groupId=([0-9a-f-]{36})/i)?.[1];
  return { id: id.toLowerCase(), groupId: group?.toLowerCase() };
}
export function parseDatasetUrl(url: string): Ref {
  const id = url.match(/(?:datasets|semanticmodels)\/([0-9a-f-]{36})/i)?.[1] ?? url.match(GUID)?.[0];
  if (!id) throw new Error("No semantic model id in this link");
  const group = url.match(/groups\/([0-9a-f-]{36})/i)?.[1];
  return { id: id.toLowerCase(), groupId: group?.toLowerCase() };
}

/* ---------------------------------- DAX ---------------------------------- */

export type Rows = Record<string, unknown>[];

export async function dax(dataset: Ref, query: string): Promise<Rows> {
  const data = await pbi(`${scope(dataset.groupId)}/datasets/${dataset.id}/executeQueries`, "POST", {
    queries: [{ query }],
    serializerSettings: { includeNulls: true },
  });
  const result = data.results[0];
  if (result.error) throw new Error(result.error.message);
  return (result.tables?.[0]?.rows ?? []).map((row: Record<string, unknown>) =>
    Object.fromEntries(Object.entries(row).map(([k, v]) => [k.replace(/^.*\[(.*)\]$/, "$1"), v])),
  );
}

/** Rows as a compact text table for the AI. */
export function rowsText(rows: Rows, max = 50) {
  if (rows.length === 0) return "(no rows)";
  const columns = Object.keys(rows[0]);
  const lines = rows.slice(0, max).map((r) => columns.map((c) => String(r[c] ?? "")).join(" | "));
  return [columns.join(" | "), ...lines, rows.length > max ? `… ${rows.length - max} more rows` : ""].filter(Boolean).join("\n");
}

/** Model structure from DAX INFO views (works with Build permission). */
export async function modelText(dataset: Ref) {
  const safe = (q: string) => dax(dataset, q).catch(() => [] as Rows);
  const [columns, measures] = await Promise.all([safe("EVALUATE INFO.VIEW.COLUMNS()"), safe("EVALUATE INFO.VIEW.MEASURES()")]);
  const tables = new Map<string, string[]>();
  const add = (table: string, item: string) => tables.set(table, [...(tables.get(table) ?? []), item]);
  for (const c of columns)
    if (!String(c.Name).startsWith("RowNumber-") && !/^(LocalDateTable|DateTableTemplate)_/.test(String(c.Table)) && !c.IsHidden)
      add(String(c.Table), `${c.Name}:${c.DataType}`);
  for (const m of measures) add(String(m.Table), `[${m.Name}]${m.Expression ? ` = ${String(m.Expression).replace(/\s+/g, " ").slice(0, 200)}` : ""}`);
  if (tables.size === 0) return "(model structure not readable)";
  return [...tables].map(([t, items]) => `'${t}': ${items.join(", ")}`).join("\n");
}

/* ------------------------------- Embedding ------------------------------- */

let client: typeof PbiClient | undefined;
let service: PbiClient.service.Service | undefined;

async function load() {
  if (!client) {
    client = await import("powerbi-client");
    await import("powerbi-report-authoring");
    service = new client.service.Service(client.factories.hpmFactory, client.factories.wpmpFactory, client.factories.routerFactory);
  }
  return { pbi: client, service: service! };
}
export const models = async () => (await load()).pbi.models;

export type EmbedMode = "view" | "edit";

export async function embedReport(element: HTMLElement, ref: Ref, mode: EmbedMode = "view") {
  const { pbi: p, service } = await load();
  service.reset(element);
  const report = service.embed(element, {
    type: "report",
    id: ref.id,
    embedUrl: `https://app.powerbi.com/reportEmbed?reportId=${ref.id}${ref.groupId ? `&groupId=${ref.groupId}` : ""}`,
    accessToken: await powerBiToken(),
    tokenType: p.models.TokenType.Aad,
    permissions: mode === "edit" ? p.models.Permissions.All : p.models.Permissions.Read,
    viewMode: mode === "edit" ? p.models.ViewMode.Edit : p.models.ViewMode.View,
    settings: {
      panes: { filters: { visible: mode === "edit" } },
      background: p.models.BackgroundType.Transparent,
      // Whole page visible by default (zoom buttons for more).
      layoutType: p.models.LayoutType.Custom,
      customLayout: { displayOption: p.models.DisplayOption.FitToPage },
    },
  }) as Report;
  await loaded(report);
  return report;
}

/** Blank report on a semantic model (native authoring UI), saved with saveAs. */
export async function embedNewReport(element: HTMLElement, dataset: Ref) {
  const { pbi: p, service } = await load();
  service.reset(element);
  const embed = service.createReport(element, {
    type: "report",
    datasetId: dataset.id,
    groupId: dataset.groupId,
    embedUrl: `https://app.powerbi.com/reportEmbed${dataset.groupId ? `?groupId=${dataset.groupId}` : ""}`,
    accessToken: await powerBiToken(),
    tokenType: p.models.TokenType.Aad,
  } as PbiClient.IEmbedConfiguration);
  await loaded(embed);
  return embed;
}

export const resetEmbed = async (element: HTMLElement) => (await load()).service.reset(element);

const loaded = (embed: PbiClient.Embed) =>
  new Promise<void>((resolve, reject) => {
    embed.on("loaded", () => resolve());
    embed.on("error", (e) => {
      const d = e.detail as { message?: string; detailedMessage?: string };
      reject(new Error(d?.detailedMessage ?? d?.message ?? "Could not load the report"));
    });
  });

/* ----------------------------- What the AI sees ---------------------------- */

const fieldName = (t: Record<string, any>) => (t.measure ? `[${t.measure}]` : `'${t.table}'[${t.column ?? t.hierarchyLevel ?? t.hierarchy}]`);
/** Active filters only ("All" / empty ones are noise for the AI). */
const active = (filters: any[]) => filters.filter((f) => (f.values?.length && f.operator !== "All") || f.conditions?.length);
const filterText = (f: any) => {
  const target = f.target ? fieldName(f.target) : "?";
  if (f.values) return `${target} ${f.operator ?? "In"} ${JSON.stringify(f.values)}`;
  if (f.conditions) return `${target} ${f.conditions.map((c: any) => `${c.operator} ${c.value ?? ""}`).join(` ${f.logicalOperator} `)}`;
  return `${target} (${f.filterType})`;
};

async function visualFields(visual: Visual) {
  const caps = await visual.getCapabilities().catch(() => null);
  const roles: string[] = [];
  for (const role of caps?.dataRoles ?? []) {
    const fields = await visual.getDataFields(role.name).catch(() => []);
    if (fields.length) roles.push(`${role.name}: ${fields.map((f) => fieldName(f as never)).join(", ")}`);
  }
  return roles.join("; ");
}

/** One visual for the AI: kind, fields, filters, slicer state, and (optionally) its data. */
export async function visualText(visual: Visual, withData: boolean) {
  const { pbi: p } = await load();
  const lines = [`- ${visual.type} "${visual.title ?? ""}" name=${visual.name}${visual.layout.displayState?.mode === 1 ? " (hidden)" : ""}`];
  const fields = await visualFields(visual);
  if (fields) lines.push(`  fields: ${fields}`);
  const filters = active(await visual.getFilters().catch(() => []));
  if (filters.length) lines.push(`  filters: ${filters.map(filterText).join("; ")}`);
  if (visual.type === "slicer") {
    const state = await visual.getSlicerState().catch(() => null);
    if (state) lines.push(`  slicer on ${state.targets?.map((t) => fieldName(t as never)).join(", ")}: ${active(state.filters).length ? active(state.filters).map(filterText).join("; ") : "nothing selected"}`);
  }
  if (withData && !["slicer", "textbox", "image", "shape", "actionButton", "basicShape"].includes(visual.type)) {
    const exported = await visual.exportData(p.models.ExportDataType.Summarized, 30).catch(() => null);
    if (exported?.data) lines.push(`  data (first rows, CSV):\n    ${exported.data.replace(/\r/g, "").trim().split("\n").slice(0, 31).join("\n    ")}`);
  }
  return lines.join("\n");
}

/** The page the user is looking at: name, filters, visuals with their data. Sent with every message. */
export async function pageText(report: Report, withData = true) {
  const pages = await report.getPages();
  const page = pages.find((p) => p.isActive) ?? pages[0];
  const visuals = await page.getVisuals();
  const [reportFilters, pageFilters] = (await Promise.all([report.getFilters().catch(() => []), page.getFilters().catch(() => [])])).map(active);
  const parts = await Promise.all(visuals.map((v) => visualText(v, withData)));
  return [
    `pages: ${pages.map((p) => `${p.displayName} (name=${p.name}${p.isActive ? ", current" : ""}${p.visibility === 1 ? ", hidden" : ""})`).join(", ")}`,
    `report filters: ${reportFilters.map(filterText).join("; ") || "none"}`,
    `current page "${page.displayName}" filters: ${pageFilters.map(filterText).join("; ") || "none"}`,
    `visuals on the current page:`,
    ...parts,
  ].join("\n");
}

/** Every page without data: the report "snapshot" stored in the library. */
export async function reportText(report: Report) {
  const out: string[] = [];
  for (const page of await report.getPages()) {
    const visuals = await page.getVisuals().catch(() => []);
    const lines = await Promise.all(visuals.map((v) => visualText(v, false)));
    out.push(`## Page "${page.displayName}" (name=${page.name}${page.visibility === 1 ? ", hidden" : ""})`, ...lines);
  }
  return out.join("\n");
}

/** Waits until the report has rendered again after a change (max 4 s). */
export const settled = (report: Report) =>
  new Promise<void>((resolve) => {
    const done = () => {
      report.off("rendered");
      resolve();
    };
    report.on("rendered", done);
    setTimeout(done, 4000);
  });

/** Basic "In" filter on a column. */
export const basicFilter = (table: string, column: string, values: unknown[]) => ({
  $schema: "http://powerbi.com/product/schema#basic",
  target: { table, column },
  operator: "In",
  values,
  filterType: 1,
});
