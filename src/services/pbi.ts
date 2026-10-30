/** Power BI: links, DAX, model structure, embedding and what the AI sees of a report page. */
import type * as PbiClient from "powerbi-client";
import { PERFORMANCE } from "@/config";
import { pbi, powerBiToken } from "./ms";

/** appId: a report opened from a Power BI app (app.powerbi.com/groups/me/apps/<appId>/reports/<id>). */
export type Ref = { id: string; groupId?: string; appId?: string };
export type Report = PbiClient.Report;
type Visual = PbiClient.VisualDescriptor;

const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
export const scope = (groupId?: string) => (groupId ? `/groups/${groupId}` : "");

/** app.powerbi.com/groups/<ws>/reports/<id>/…, …/apps/<appId>/reports/<id>/… (or a bare id) → ids. */
export function parseReportUrl(url: string): Ref {
  const id = url.match(/reports\/([0-9a-f-]{36})/i)?.[1] ?? url.match(/reportId=([0-9a-f-]{36})/i)?.[1] ?? url.match(GUID)?.[0];
  if (!id) throw new Error("No report id in this link");
  const app = url.match(/apps\/([0-9a-f-]{36})/i)?.[1] ?? url.match(/appId=([0-9a-f-]{36})/i)?.[1];
  if (app) return { id: id.toLowerCase(), appId: app.toLowerCase() };
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

/** A DAX error with what to do about it (weak models otherwise guess another name, or a wrong cause). */
function daxError(message: string) {
  const hint = /cannot be found|can't be found|not found|introuvable|n'existe pas|cannot find|cannot be determined|doesn't exist/i.test(message)
    ? " → A name does not exist: look up the exact table / column / measure names with search_model or describe_table (never guess or translate them), then fix the query."
    : /^Query \(\d+, \d+\)|syntax/i.test(message)
      ? " → (line, column) of the query: fix the query exactly there."
      : "";
  return new Error(`${message}${hint}`);
}

/* Limits that keep the models fast for everyone (docs/performance.md): PERFORMANCE in src/config.ts. */

/** The rows of a query that were cut at the row limit (more rows exist). */
export type Truncated = Rows & { truncated?: number };

/** The `ORDER BY a DESC, b` items of a query as TOPN arguments ("a, DESC, b, ASC"). */
const orderArgs = (orderBy: string) =>
  splitTop(orderBy)
    .map((item) => {
      const m = item.trim().match(/^([\s\S]*?)\s+(ASC|DESC)$/i);
      return m ? `${m[1]}, ${m[2].toUpperCase()}` : `${item.trim()}, ASC`;
    })
    .join(", ");

/** Splits at the commas outside parentheses, brackets and quotes. */
function splitTop(text: string) {
  const parts: string[] = [];
  let depth = 0;
  let quote = "";
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quote) quote = c === quote ? "" : quote;
    else if (c === '"' || c === "'") quote = c;
    else if (c === "(" || c === "[") depth++;
    else if (c === ")" || c === "]") depth--;
    else if (c === "," && depth === 0) (parts.push(text.slice(from, i)), (from = i + 1));
  }
  return [...parts, text.slice(from)];
}

/**
 * The query with its last EVALUATE limited to `limit` rows (TOPN, the ORDER BY kept as its order):
 * the model computes no more rows than the app shows. Undefined when the query has another shape.
 */
export function limitRows(query: string, limit: number) {
  const at = [...query.matchAll(/\bEVALUATE\b/gi)].at(-1)?.index;
  if (at === undefined) return undefined;
  const rest = query.slice(at + "EVALUATE".length);
  const order = rest.search(/\bORDER\s+BY\b/i);
  if (/\bSTART\s+AT\b/i.test(rest)) return undefined;
  const body = (order < 0 ? rest : rest.slice(0, order)).trim().replace(/;$/, "");
  if (!body || /^ROW\s*\(/i.test(body) || /^\{/.test(body)) return undefined;
  const orderBy = order < 0 ? "" : rest.slice(order).replace(/^ORDER\s+BY\b/i, "").trim().replace(/;$/, "");
  return `${query.slice(0, at)}EVALUATE\nTOPN(${limit}, ${body}${orderBy ? `, ${orderArgs(orderBy)}` : ""})${orderBy ? `\nORDER BY ${orderBy}` : ""}`;
}

/** At most `concurrency` queries at a time per model, `intervalMs` apart (the model's capacity is shared). */
const queues = new Map<string, { running: number; last: number; waiting: (() => void)[] }>();
async function slot(model: string) {
  const q = queues.get(model) ?? queues.set(model, { running: 0, last: 0, waiting: [] }).get(model)!;
  while (q.running >= Math.max(1, PERFORMANCE.concurrency)) await new Promise<void>((r) => q.waiting.push(r));
  q.running++;
  const wait = q.last + PERFORMANCE.intervalMs - Date.now();
  if (wait > 0) await new Promise((r) => setTimeout(r, wait));
  q.last = Date.now();
  return () => {
    q.running--;
    q.waiting.shift()?.();
  };
}

/** The same query on the same model within `cacheMs`: the rows already read. */
const cache = new Map<string, { at: number; rows: Promise<Truncated> }>();

/**
 * Runs a DAX query: at most `maxRows` rows (more are cut, `truncated` says how many were asked for at
 * least), stopped after `timeoutMs`, queued per model and cached a little (src/config.ts PERFORMANCE).
 */
export async function dax(dataset: Ref, query: string): Promise<Truncated> {
  const key = `${dataset.id}\n${query}`;
  const hit = cache.get(key);
  if (hit && Date.now() - hit.at < PERFORMANCE.cacheMs) return hit.rows;
  const rows = runDax(dataset, query);
  cache.set(key, { at: Date.now(), rows });
  rows.catch(() => cache.delete(key));
  for (const [k, v] of cache) if (Date.now() - v.at > PERFORMANCE.cacheMs) cache.delete(k);
  return rows;
}

async function runDax(dataset: Ref, query: string): Promise<Truncated> {
  const max = Math.max(1, PERFORMANCE.maxRows);
  const limited = limitRows(query, max + 1);
  const done = await slot(dataset.id);
  try {
    let rows: Rows;
    try {
      rows = await executeDax(dataset, limited ?? query);
    } catch (e) {
      // The limit made the query invalid (an unusual shape): asked again as written, cut here.
      if (!limited || /took longer/.test(String(e))) throw e;
      rows = await executeDax(dataset, query);
    }
    const out: Truncated = rows.slice(0, max);
    if (rows.length > max) out.truncated = rows.length;
    return out;
  } finally {
    done();
  }
}

async function executeDax(dataset: Ref, query: string): Promise<Rows> {
  const abort = new AbortController();
  const timer = setTimeout(() => abort.abort(), PERFORMANCE.timeoutMs);
  const data = await pbi(`${scope(dataset.groupId)}/datasets/${dataset.id}/executeQueries`, "POST", { queries: [{ query }], serializerSettings: { includeNulls: true } }, abort.signal)
    .catch((e) => {
      if (abort.signal.aborted) throw new Error(`The query took longer than ${PERFORMANCE.timeoutMs / 1000} s and was stopped (limit of the app, to keep the model fast). Make it lighter: fewer columns, a filter, SUMMARIZECOLUMNS rather than FILTER on a whole table, or count the rows first.`);
      throw daxError(e instanceof Error ? e.message : String(e));
    })
    .finally(() => clearTimeout(timer));
  const result = data.results[0];
  if (result.error) throw daxError(result.error.message);
  return (result.tables?.[0]?.rows ?? []).map((row: Record<string, unknown>) =>
    Object.fromEntries(Object.entries(row).map(([k, v]) => [k.replace(/^.*\[(.*)\]$/, "$1"), v])),
  );
}

/** Refreshes on a Pro (shared) capacity: 8 a day per model, by API and scheduled together (refreshes started in Power BI itself are not counted). */
export const REFRESH_QUOTA = 8;

/** The refreshes that count in the daily quota (by API or scheduled) over the last 24 hours. */
export async function refreshesToday(dataset: Ref) {
  const { value } = await pbi<{ value: { refreshType: string; startTime: string }[] }>(`${scope(dataset.groupId)}/datasets/${dataset.id}/refreshes?$top=30`);
  const since = Date.now() - 24 * 3600 * 1000;
  return value.filter((r) => (r.refreshType === "ViaApi" || r.refreshType === "Scheduled") && Date.parse(r.startTime) >= since).length;
}

/* ----------------------------- Column values ----------------------------- */

const columnRef = (table: string, column: string) => `'${table.replace(/'/g, "''")}'[${column.replace(/]/g, "]]")}]`;
const literal = (v: unknown) => (typeof v === "number" ? String(v) : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : `"${String(v).replace(/"/g, '""')}"`);

/** Distinct values of a column (first `VALUES_SHOWN`), kept for the session: what slicers and filters can take. */
const VALUES_SHOWN = 40;
const valuesCache = new Map<string, Promise<{ values: unknown[]; more: boolean }>>();
export function columnValues(dataset: Ref, table: string, column: string) {
  const key = `${dataset.id}|${table}|${column}`;
  if (!valuesCache.has(key)) {
    const run = dax(dataset, `EVALUATE TOPN(${VALUES_SHOWN + 1}, VALUES(${columnRef(table, column)}))`).then((rows) => {
      const values = rows.map((r) => Object.values(r)[0]).filter((v) => v !== null && v !== "");
      return { values: values.slice(0, VALUES_SHOWN), more: values.length > VALUES_SHOWN };
    });
    valuesCache.set(key, run);
    run.catch(() => valuesCache.delete(key));
  }
  return valuesCache.get(key)!;
}

/**
 * The values to filter on, as the model holds them: matched without case ("fall" → "Fall"), and
 * refused when they do not exist (the error lists the valid ones), so the AI never filters on a value
 * that empties the visuals. Without access to the model (no Build right), kept as given.
 */
export async function checkValues(dataset: Ref | undefined, table: string, column: string, given: unknown[]) {
  if (!dataset || !given.length) return given;
  const found = await dax(dataset, `EVALUATE FILTER(VALUES(${columnRef(table, column)}), ${columnRef(table, column)} IN {${given.map(literal).join(", ")}})`)
    .then((rows) => rows.map((r) => Object.values(r)[0]))
    .catch(() => undefined);
  if (!found) return given;
  const same = (a: unknown, b: unknown) => String(a).toLowerCase() === String(b).toLowerCase();
  const missing = given.filter((g) => !found.some((f) => same(f, g)));
  if (missing.length) {
    const { values, more } = await columnValues(dataset, table, column).catch(() => ({ values: [] as unknown[], more: false }));
    throw new Error(
      `${table}[${column}] has no value ${missing.map((m) => JSON.stringify(m)).join(", ")}. ${values.length ? `Its values${more ? " (first " + values.length + ")" : ""}: ${values.map((v) => JSON.stringify(v)).join(", ")}.` : ""} Pick the matching one (e.g. a translation or synonym of the user's word) and retry${more ? ", or search them with run_dax" : ""}.`,
    );
  }
  return given.map((g) => found.find((f) => same(f, g)) ?? g);
}

/** Rows as a compact text table for the AI. */
export function rowsText(rows: Truncated, max = 50) {
  if (rows.length === 0) return "(no rows)";
  const cut = rows.truncated ? `\n⚠ Cut at the app's limit of ${rows.length} rows (the query returns more): aggregate, filter or TOPN; to know how many, EVALUATE ROW("n", COUNTROWS(<the same table>)).` : "";
  const columns = Object.keys(rows[0]);
  // Long values (e.g. base64 images) shown by their start and length: one would fill the AI's context.
  const cell = (v: unknown) => {
    const text = String(v ?? "");
    // 32,766 = Power BI's limit for a text value: a value of exactly that length was cut.
    return text.length > 120 ? `${text.slice(0, 60)}… (${text.length} characters${text.length === 32766 ? ", CUT at Power BI's limit" : ""})` : text;
  };
  const lines = rows.slice(0, max).map((r) => columns.map((c) => cell(r[c])).join(" | "));
  return [columns.join(" | "), ...lines, rows.length > max ? `… ${rows.length - max} more rows` : ""].filter(Boolean).join("\n") + cut;
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

const embedUrlOf = (ref: Ref) => `https://app.powerbi.com/reportEmbed?reportId=${ref.id}${ref.appId ? `&appId=${ref.appId}` : ref.groupId ? `&groupId=${ref.groupId}` : ""}`;

/** `bare`: the page alone — no filter pane, no page tabs (slides crop zones of it). */
export async function embedReport(element: HTMLElement, ref: Ref, mode: EmbedMode = "view", { bare = false } = {}) {
  const { pbi: p, service } = await load();
  service.reset(element);
  const report = service.embed(element, {
    type: "report",
    id: ref.id,
    embedUrl: embedUrlOf(ref),
    accessToken: await powerBiToken(),
    tokenType: p.models.TokenType.Aad,
    permissions: mode === "edit" ? p.models.Permissions.All : p.models.Permissions.Read,
    viewMode: mode === "edit" ? p.models.ViewMode.Edit : p.models.ViewMode.View,
    settings: {
      // The filter pane is there in both modes, folded when viewing (the user opens it if needed).
      panes: bare ? { filters: { visible: false }, pageNavigation: { visible: false } } : { filters: { visible: true, expanded: mode === "edit" } },
      background: p.models.BackgroundType.Transparent,
      // Whole page visible by default (zoom buttons for more).
      layoutType: p.models.LayoutType.Custom,
      customLayout: { displayOption: p.models.DisplayOption.FitToPage },
    },
  }) as Report;
  await loaded(report);
  return report;
}

/**
 * One visual of a report, alone (guide preview), with the main report's filters (`filters`: report
 * level, and page level — the slicers' selections go there). Not a bookmark: applied to a visual
 * embed, it brought the rest of the page back over the visual. Embedding again in the same element
 * reuses its frame: only the first one is slow.
 */
export async function embedVisual(element: HTMLElement, ref: Ref, page: string, visual: string, filters?: { report: unknown[]; page: unknown[] }) {
  const { pbi: p, service } = await load();
  const embed = service.embed(element, {
    type: "visual",
    id: ref.id,
    embedUrl: embedUrlOf(ref),
    accessToken: await powerBiToken(),
    tokenType: p.models.TokenType.Aad,
    pageName: page,
    visualName: visual,
    settings: { panes: { filters: { visible: false } }, background: p.models.BackgroundType.Transparent },
  } as PbiClient.IVisualEmbedConfiguration) as PbiClient.Visual;
  await loaded(embed);
  if (filters) await setVisualFilters(embed, filters);
  return embed;
}

/** Report and page filters of a visual embed (kept in step with the main report). */
export async function setVisualFilters(embed: PbiClient.Visual, filters: { report: unknown[]; page: unknown[] }) {
  const { pbi: p } = await load();
  await embed.setFilters(filters.report as never, p.models.FiltersLevel.Report).catch(() => undefined);
  await embed.setFilters(filters.page as never, p.models.FiltersLevel.Page).catch(() => undefined);
}

export const resetEmbed = async (element: HTMLElement) => (await load()).service.reset(element);

/** Next "loaded" (or error) of an embed; earlier waits are dropped (an embed can be loaded again). */
const loaded = (embed: PbiClient.Embed) =>
  new Promise<void>((resolve, reject) => {
    embed.off("loaded");
    embed.off("error");
    embed.on("loaded", () => resolve());
    embed.on("error", (e) => {
      const d = e.detail as { message?: string; detailedMessage?: string };
      reject(new Error(d?.detailedMessage ?? d?.message ?? "Could not load the report"));
    });
  });

/* ----------------------------- What the AI sees ---------------------------- */

/** A field as DAX writes it; a field the SDK did not describe (it happens) is "?" rather than a crash. */
const fieldName = (t?: Record<string, any> | null) => (!t ? "?" : t.measure ? `[${t.measure}]` : `'${t.table}'[${t.column ?? t.hierarchyLevel ?? t.hierarchy}]`);
/** Active filters only ("All" / empty ones are noise for the AI). */
const active = (filters: any[]) => filters.filter((f) => (f.values?.length && f.operator !== "All") || f.conditions?.length || f.hierarchyData?.length);
/** Selected paths of a hierarchy filter ("2024 > Jan"): a node is selected whole, or through its children. */
const paths = (nodes: any[] = [], prefix: unknown[] = []): string[] =>
  nodes.flatMap((n) => (n.operator === "Selected" ? [[...prefix, n.value].join(" > ")] : n.children ? paths(n.children, [...prefix, n.value]) : []));
const filterText = (f: any) => {
  if (f.hierarchyData) return `${(Array.isArray(f.target) ? f.target : [f.target]).map(fieldName).join(" > ")} = ${JSON.stringify(paths(f.hierarchyData))}`;
  const target = f.target ? fieldName(f.target) : "?";
  if (f.values) return `${target} ${f.operator ?? "In"} ${JSON.stringify(f.values)}`;
  if (f.conditions) return `${target} ${f.conditions.map((c: any) => `${c.operator} ${c.value ?? ""}`).join(` ${f.logicalOperator} `)}`;
  return `${target} (${f.filterType})`;
};

async function visualFields(visual: Visual) {
  // Methods of the authoring SDK: missing if it could not load.
  if (typeof visual.getCapabilities !== "function") return "";
  const caps = await visual.getCapabilities().catch(() => null);
  const roles: string[] = [];
  for (const role of caps?.dataRoles ?? []) {
    const fields = await visual.getDataFields(role.name).catch(() => []);
    if (fields.length) roles.push(`${role.name}: ${fields.map((f) => fieldName(f as never)).join(", ")}`);
  }
  return roles.join("; ");
}

const sameTitle = (a = "", b = "") => a.toLowerCase().replace(/[\s_-]+/g, " ").trim() === b.toLowerCase().replace(/[\s_-]+/g, " ").trim();

/**
 * A visual by its name (its id) or its title. A title shared by several visuals is refused: the AI
 * must then give the name, else it would act on the wrong one.
 */
export function findVisual<V extends { name: string; title?: string }>(visuals: V[], nameOrTitle: string): V {
  const byName = visuals.find((v) => v.name === nameOrTitle);
  if (byName) return byName;
  const matches = visuals.filter((v) => sameTitle(v.title, nameOrTitle));
  if (matches.length > 1) throw new Error(`${matches.length} visuals are titled "${nameOrTitle}": give the name (id) of the one you mean: ${matches.map((v) => `name=${v.name}`).join(", ")}.`);
  if (!matches.length) throw new Error(`Visual "${nameOrTitle}" not on the current page. Visuals: ${visuals.map((v) => `"${v.title ?? ""}" (name=${v.name})`).join(", ")}. Use the name.`);
  return matches[0];
}

/**
 * One visual for the AI: kind, fields, filters, slicer state (with the values it offers, when the
 * model can be queried: `dataset`), and (optionally) its data. `shared`: its title is not unique.
 */
export async function visualText(visual: Visual, withData: boolean, shared = false, dataset?: Ref) {
  const { pbi: p } = await load();
  const { x = 0, y = 0, width = 0, height = 0 } = visual.layout;
  const box = `x=${Math.round(x)} y=${Math.round(y)} w=${Math.round(width)} h=${Math.round(height)} z=${Math.round(visual.layout.z ?? 0)}`;
  const lines = [`- ${visual.type} "${visual.title ?? ""}" name=${visual.name} ${box}${visual.layout.displayState?.mode === 1 ? " (hidden)" : ""}${shared ? " ⚠ SAME TITLE AS ANOTHER VISUAL: refer to it ONLY by name=" + visual.name : ""}`];
  const fields = await visualFields(visual);
  if (fields) lines.push(`  fields: ${fields}`);
  const filters = active(await visual.getFilters().catch(() => []));
  if (filters.length) lines.push(`  filters: ${filters.map(filterText).join("; ")}`);
  if (visual.type === "slicer") {
    const state = await visual.getSlicerState().catch(() => null);
    if (state) lines.push(`  slicer on ${state.targets?.map((t) => fieldName(t as never)).join(", ")}: ${active(state.filters).length ? active(state.filters).map(filterText).join("; ") : "nothing selected"}`);
    const levels = (state?.targets ?? []) as { table?: string; column?: string; hierarchyLevel?: string }[];
    if (levels.length > 1) lines.push(`  hierarchy slicer (${levels.length} levels): select with paths "top > sub" (e.g. "${levels.map((_, i) => `level${i + 1}`).join(" > ")}"); a top value alone selects it whole`);
    // Options of the first (or only) level.
    const t = levels[0];
    const column = t?.column ?? t?.hierarchyLevel;
    const options = dataset && t?.table && column ? await columnValues(dataset, t.table, column).catch(() => undefined) : undefined;
    if (options?.values.length) lines.push(`  options${levels.length > 1 ? " (top level)" : ""}${options.more ? ` (first ${options.values.length})` : ""}: ${options.values.map((v) => JSON.stringify(v)).join(", ")}`);
  }
  if (withData && !["slicer", "textbox", "image", "shape", "actionButton", "basicShape"].includes(visual.type)) {
    const exported = await visual.exportData(p.models.ExportDataType.Summarized, 30).catch(() => null);
    if (exported?.data) lines.push(`  data (first rows, CSV):\n    ${exported.data.replace(/\r/g, "").trim().split("\n").slice(0, 31).join("\n    ")}`);
  }
  return lines.join("\n");
}

/** The page the user is looking at: name, filters, visuals with their data (and slicer options: `dataset`). */
export async function pageText(report: Report, withData = true, dataset?: Ref, progress?: (done: number, total: number) => void) {
  const pages = await report.getPages();
  const page = pages.find((p) => p.isActive) ?? pages[0];
  const visuals = await page.getVisuals();
  const [reportFilters, pageFilters] = (await Promise.all([report.getFilters().catch(() => []), page.getFilters().catch(() => [])])).map(active);
  const shared = (v: Visual) => visuals.filter((x) => sameTitle(x.title, v.title)).length > 1 && !!v.title;
  // Exporting each visual's data is what takes time: reported visual by visual.
  let done = 0;
  progress?.(0, visuals.length);
  const parts = await Promise.all(visuals.map((v) => visualText(v, withData, shared(v), dataset).finally(() => progress?.(++done, visuals.length))));
  return [
    `pages: ${pages.map((p) => `${p.displayName} (name=${p.name}${p.isActive ? ", current" : ""}${p.visibility === 1 ? ", hidden" : ""})`).join(", ")}`,
    `report filters: ${reportFilters.map(filterText).join("; ") || "none"}`,
    `current page "${page.displayName}" filters: ${pageFilters.map(filterText).join("; ") || "none"}`,
    `visuals on the current page (page ${page.defaultSize?.width ?? 1280}×${page.defaultSize?.height ?? 720}, z = stacking order, higher is in front):`,
    ...parts,
    ...layoutIssues(visuals, page.defaultSize),
  ].join("\n");
}

const BACKGROUNDS = ["shape", "basicShape", "image"];

/** Overlapping visuals and visuals outside the page, so the AI sees layout mistakes. */
export function layoutIssues(visuals: Visual[], size?: { width?: number; height?: number }) {
  const W = size?.width ?? 1280, H = size?.height ?? 720;
  const shown = visuals.filter((v) => v.layout.displayState?.mode !== 1);
  const box = (v: Visual) => ({ x: v.layout.x ?? 0, y: v.layout.y ?? 0, r: (v.layout.x ?? 0) + (v.layout.width ?? 0), b: (v.layout.y ?? 0) + (v.layout.height ?? 0) });
  const label = (v: Visual) => `"${v.title ?? v.type}" (${v.name})`;
  const issues: string[] = [];
  for (const v of shown) {
    const a = box(v);
    if (a.x < 0 || a.y < 0 || a.r > W + 1 || a.b > H + 1) issues.push(`- ${label(v)} goes outside the page (right=${Math.round(a.r)} bottom=${Math.round(a.b)})`);
  }
  const front = shown.filter((v) => !BACKGROUNDS.includes(v.type));
  for (let i = 0; i < front.length; i++)
    for (let j = i + 1; j < front.length; j++) {
      const a = box(front[i]), b = box(front[j]);
      const w = Math.min(a.r, b.r) - Math.max(a.x, b.x), h = Math.min(a.b, b.b) - Math.max(a.y, b.y);
      if (w > 1 && h > 1) issues.push(`- ${label(front[i])} and ${label(front[j])} overlap (${Math.round(w)}×${Math.round(h)})`);
    }
  return issues.length ? ["layout problems (fix them when you change the layout):", ...issues] : ["layout: no overlap, everything inside the page"];
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
