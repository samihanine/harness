/**
 * Report editor: the AI changes the report's definition (PBIR) — report, theme, pages, visuals, report
 * measures — in a draft saved when its answer is done (see `finish` in the editor). New visuals come
 * from a short spec (visual-spec.ts); everything else is changed with targeted JSON Patch operations.
 */
import { z } from "zod";
import { dax, findVisual, rowsText, visualText, type Ref, type Report } from "@/lib/pbi";
import type { ModelInfo } from "@/lib/model";
import { applyOp, deleteDraftFiles, discardDraft, draftFiles, patchDraftFile, pendingFiles, putDraftFile, readDraftFile, readDraftJson, undoLastSave, type PatchOp } from "@/lib/report-draft";
import { ROLES, TYPE_NOTES, roleKey, buildPage, buildVisual, formatOps, formatOptions, ignoredText, newName, pageFormatOps, topNFilter, type Format, type ReportMeasures } from "@/lib/visual-spec";
import { tool, type Tool } from "./loop";
import { queryTool } from "./pbi-tools";
import { LAYOUT_HELP, layoutProblems, layoutRows } from "@/lib/layout";
import type { DefinedMeasure } from "@/lib/query-spec";
import { pageCopy } from "@/lib/report-copy";
import type { ReportSchema } from "@/lib/report-schema";
import { buttonVisual, conditionalOps, linkProperties, referenceLineOps, type Action } from "@/lib/visual-extras";

const P = "definition/pages";
const REPORT = "definition/report.json";
const PAGES = `${P}/pages.json`;
const EXTENSION = "definition/reportExtensions.json";

/** How the definition is written: given to the AI once (its first message). */
export const PBIR_GUIDE = `Report definition (PBIR JSON), the files you change:
- report.json: themeCollection, filterConfig.filters (report filters), objects (e.g. outspacePane = filter pane), settings.
- pages/pages.json: pageOrder (page names), activePageName.
- pages/<page>/page.json: name, displayName, width, height, displayOption, objects.background / objects.outspace (page color), filterConfig, visualInteractions [{source, target, type: DataFilter | HighlightFilter | NoFilter}], visibility ("HiddenInViewMode").
- pages/<page>/visuals/<visual>/visual.json: name, position {x, y, z, width, height, tabOrder}, isHidden, visual: {visualType, query.queryState.<Role>.projections [{field, queryRef, nativeQueryRef}], query.sortDefinition, objects (chart parts: legend, categoryAxis, valueAxis, labels, dataPoint, columnHeaders, values…), visualContainerObjects (title, subTitle, background, border, dropShadow, padding)}.
- reportExtensions.json: report measures (use update_measure / delete_measure).
- the custom theme (get_theme / update_theme): dataColors, background, foreground, textClasses, visualStyles {<visualType or "*">: {"*": {<object>: [{<property>: value}]}}} — the best place for a style shared by every visual.
In visual.json (update_visual only), values in objects are expressions: {"expr": {"Literal": {"Value": "'text'"}}} (text in single quotes), "12D" (number), "true" / "false"; a color: {"solid": {"color": {"expr": {"Literal": {"Value": "'#1F77B4'"}}}}}. Each object is a list of {"properties": {…}} (per series / per item: add a "selector"). format_visual and create_visual's "format" write these for you from plain values.
Fields in a query: {"Column": {"Expression": {"SourceRef": {"Entity": "Table"}}, "Property": "Column"}}, {"Measure": …} (report measures: SourceRef {"Schema": "extension", "Entity": "Table"}), {"Aggregation": {"Expression": {"Column": …}, "Function": 0 sum | 1 avg | 2 count | 3 min | 4 max}}. To add or change fields, prefer create_visual / set_visual_fields: they write this for you and check the names.`;

type Ctx = {
  report: () => Report | undefined;
  ref: () => Ref | undefined;
  dataset: () => Ref | undefined;
  info: () => ModelInfo | undefined;
  /** The reference reports chosen by the user (Builder › Reports), with their schemas. */
  references?: () => Promise<{ name: string; schema?: ReportSchema }[]>;
};

/** Where a visual given without a place goes: under the page's other visuals (layout_page then arranges them). */
function freeSpot(others: { json: any }[], page: { width?: number; height?: number }) {
  const bottom = Math.max(0, ...others.map((v) => Number(v.json.position?.y ?? 0) + Number(v.json.position?.height ?? 0)));
  const height = 240;
  const y = Math.min(bottom + 16, Math.max(0, (page.height ?? 720) - height));
  return { x: 20, y, width: Math.min(420, (page.width ?? 1280) - 40), height };
}

/**
 * What is wrong in the draft, found by code: visuals overlapping or out of their page, report measures
 * whose DAX fails. Empty when all is fine.
 */
export async function draftProblems(ctx: Ctx, save: () => Promise<unknown>, pages?: string[]) {
  const ref = ctx.ref();
  if (!ref) return [];
  const files = await draftFiles(ref, save);
  const problems: string[] = [];
  for (const p of await pagesOf(files)) {
    if (pages && !pages.includes(p.name)) continue;
    const visuals = visualsOf(files, p.name).filter((v) => !v.json.parentGroupName).map((v) => ({ name: v.name, hidden: !!v.json.isHidden, x: Number(v.json.position?.x ?? 0), y: Number(v.json.position?.y ?? 0), width: Number(v.json.position?.width ?? 0), height: Number(v.json.position?.height ?? 0) }));
    for (const x of layoutProblems(visuals, { width: p.json.width ?? 1280, height: p.json.height ?? 720 })) problems.push(`page "${p.json.displayName}": ${x}`);
  }
  const dataset = ctx.dataset();
  const measures = (measuresOf(files) as (ReportMeasures[number] & { expression?: string })[]).filter((m) => m.expression);
  if (dataset && measures.length) {
    const defs = measures.map((m) => `MEASURE '${m.table.replace(/'/g, "''")}'[${m.name}] = ${m.expression}`).join("\n");
    for (const m of measures) await dax(dataset, `DEFINE\n${defs}\nEVALUATE ROW("v", [${m.name}])`).catch((e) => problems.push(`report measure [${m.name}] fails: ${e instanceof Error ? e.message.split("\n")[0].slice(0, 200) : e}`));
  }
  return problems;
}

const ops = z.array(z.object({ op: z.enum(["add", "replace", "remove"]), path: z.string().describe("JSON Pointer, e.g. /visual/objects/legend/0/properties/show"), value: z.unknown().optional() }));

/** A visual's title as PBIR stores it (a text literal), when it has one. */
const titleOf = (v: any) => String(v?.visual?.visualContainerObjects?.title?.[0]?.properties?.text?.expr?.Literal?.Value ?? "").replace(/^'(.*)'$/s, "$1").replace(/''/g, "'") || undefined;

/** Pages of the draft in their order: name, display name, size. */
async function pagesOf(files: Map<string, unknown>) {
  const order: string[] = (files.get(PAGES) as any)?.pageOrder ?? [];
  const names = [...new Set([...order, ...[...files.keys()].map((f) => f.match(/^definition\/pages\/([^/]+)\/page\.json$/)?.[1]).filter(Boolean) as string[]])];
  return names.map((name) => ({ name, json: files.get(`${P}/${name}/page.json`) as any })).filter((p) => p.json);
}

/** Visuals of a page in the draft. */
const visualsOf = (files: Map<string, unknown>, page: string) =>
  [...files.entries()]
    .filter(([f]) => f.startsWith(`${P}/${page}/visuals/`) && f.endsWith("/visual.json"))
    .map(([file, json]: [string, any]) => ({ file, name: String(json.name), title: titleOf(json), json }));

/** Report measures of the draft (reportExtensions.json). */
const measuresOf = (files: Map<string, unknown>): ReportMeasures => ((files.get(EXTENSION) as any)?.entities ?? []).flatMap((e: any) => (e.measures ?? []).map((m: any) => ({ table: e.name, name: m.name, expression: m.expression })));

/** The lists the AI gets with every message: pages, the current page's visuals, report measures, pending changes. */
export async function editorState(ctx: Ctx, save: () => Promise<unknown>) {
  const ref = ctx.ref();
  if (!ref) return "No report open.";
  const files = await draftFiles(ref, save);
  const pages = await pagesOf(files);
  const live = await ctx.report()?.getPages().catch(() => undefined);
  const current = live?.find((p) => p.isActive)?.name ?? pages[0]?.name;
  const page = pages.find((p) => p.name === current);
  const visuals = page ? visualsOf(files, page.name) : [];
  const short = (v: (typeof visuals)[number]) => {
    const { x = 0, y = 0, width = 0, height = 0, z = 0 } = v.json.position ?? {};
    const roles = Object.entries<any>(v.json.visual?.query?.queryState ?? {}).map(([role, s]) => `${role}: ${(s.projections ?? []).map((p: any) => p.queryRef).join(", ")}`);
    return `- ${v.json.visual?.visualType ?? "group"} name=${v.name}${v.title ? ` "${v.title}"` : ""} x=${Math.round(x)} y=${Math.round(y)} w=${Math.round(width)} h=${Math.round(height)} z=${z}${v.json.isHidden ? " hidden" : ""}${roles.length ? ` | ${roles.join("; ")}` : ""}`;
  };
  const measures = measuresOf(files) as (ReportMeasures[number] & { expression?: string })[];
  const pending = pendingFiles(ref);
  return [
    `Pages (list_pages): ${pages.map((p) => `"${p.json.displayName}" (name=${p.name}, ${p.json.width}×${p.json.height}${p.name === current ? ", CURRENT" : ""}${p.json.visibility === "HiddenInViewMode" ? ", hidden" : ""}, ${visualsOf(files, p.name).length} visuals)`).join("; ") || "none"}`,
    page ? `Visuals of the current page "${page.json.displayName}" (list_visuals; z = stacking order):\n${visuals.map(short).join("\n") || "none"}` : "",
    `Report measures (list_measures): ${measures.map((m) => `'${m.table}'[${m.name}] = ${String(m.expression ?? "").replace(/\s+/g, " ").slice(0, 120)}`).join("; ") || "none"}`,
    pending.length ? `Unsaved changes in your draft (saved at the end of your answer): ${pending.join(", ")}` : "",
  ]
    .filter(Boolean)
    .join("\n\n");
}

export function editorTools(ctx: Ctx, save: () => Promise<unknown>): Tool[] {
  const ref = () => {
    const r = ctx.ref();
    if (!r) throw new Error("Open a report first.");
    return r;
  };
  const files = () => draftFiles(ref(), save);
  /** A page by name or display name; none given = the current page of the editor. */
  const pageOf = async (given?: string) => {
    const pages = await pagesOf(await files());
    const current = (await ctx.report()?.getPages().catch(() => undefined))?.find((p) => p.isActive)?.name;
    const name = given?.trim() || current || pages[0]?.name;
    const page = pages.find((p) => p.name === name || p.json.displayName?.toLowerCase() === name?.toLowerCase());
    if (!page) throw new Error(`No page "${given}". Pages: ${pages.map((p) => `"${p.json.displayName}" (${p.name})`).join(", ")}.`);
    return page;
  };
  const visualOf = async (visual: string, page?: string) => {
    const p = await pageOf(page);
    return { page: p, visual: findVisual(visualsOf(await files(), p.name), visual) };
  };
  const patch = (file: string, operations: PatchOp[]) => patchDraftFile(ref(), save, file, operations);
  const read = (file: string) => readDraftFile(ref(), save, file);
  const pageArg = z.string().optional().describe("Page name or display name (default: the current page)");
  const visualArg = z.string().describe("Visual name (id), or its title when no other visual has it");
  const formatArg = z.record(z.string(), z.record(z.string(), z.unknown())).describe('By object, plain values: {"title": {"fontSize": 16, "fontColor": "#203247"}, "labels": {"show": true}}');
  /** The custom theme file, made when the report has none (registered in report.json). */
  const themeFile = async () => {
    const f = await files();
    const report = f.get(REPORT) as any;
    const custom = report?.themeCollection?.customTheme?.name;
    const item = report?.resourcePackages?.flatMap((p: any) => p.items ?? []).find((i: any) => i.name === custom && i.type === "CustomTheme");
    if (item) {
      // Power BI lists the theme without ".json" once it saved the report; the file keeps it.
      const path = `StaticResources/RegisteredResources/${item.path}`;
      return f.has(path) || !f.has(`${path}.json`) ? path : `${path}.json`;
    }
    const name = `HubTheme${newName().slice(0, 6)}`;
    const next = structuredClone(report);
    next.themeCollection = { ...next.themeCollection, customTheme: { name, reportVersionAtImport: next.themeCollection?.baseTheme?.reportVersionAtImport, type: "RegisteredResources" } };
    const packages = next.resourcePackages ?? [];
    const registered = packages.find((p: any) => p.type === "RegisteredResources");
    const entry = { name, path: `${name}.json`, type: "CustomTheme" };
    if (registered) registered.items = [...(registered.items ?? []), entry];
    else packages.push({ name: "RegisteredResources", type: "RegisteredResources", items: [entry] });
    next.resourcePackages = packages;
    await putDraftFile(ref(), save, REPORT, next);
    await putDraftFile(ref(), save, `StaticResources/RegisteredResources/${name}.json`, { name });
    return `StaticResources/RegisteredResources/${name}.json`;
  };
  const writeVisual = async (page: string, spec: Parameters<typeof buildVisual>[0], format?: Format, name = newName()) => {
    const f = await files();
    const others = visualsOf(f, page);
    const z = Math.max(0, ...others.map((v) => Number(v.json.position?.z ?? 0))) + 1000;
    const json = buildVisual(spec, name, z, ctx.info(), measuresOf(f), others[0]?.json.$schema);
    // Formatted before it is written: a wrong value refuses the whole visual (nothing half made);
    // names that do not exist on this type are skipped and said.
    const formatted = format ? formatOps(json, format) : { ops: [], ignored: [] };
    formatted.ops.forEach((op) => applyOp(json, op));
    await putDraftFile(ref(), save, `${P}/${page}/visuals/${name}/visual.json`, json);
    return { name, ignored: formatted.ignored };
  };

  return [
    tool({
      name: "get_report",
      description: "Reads report.json (theme references, report filters, settings, report-level formatting).",
      args: z.object({}),
      readOnly: true,
      run: async () => read(REPORT),
    }),
    tool({
      name: "update_report",
      description: "Changes report.json with JSON Patch operations (e.g. report filters, settings).",
      args: z.object({ operations: ops }),
      run: async ({ operations }) => patch(REPORT, operations as PatchOp[]),
    }),
    tool({
      name: "get_theme",
      description: "Reads the report's custom theme (colors, text classes, default formatting of every visual type).",
      args: z.object({}),
      readOnly: true,
      run: async () => {
        const file = await themeFile();
        return read(file);
      },
    }),
    tool({
      name: "update_theme",
      description: 'Changes the custom theme with JSON Patch operations (made when the report has none), e.g. add /dataColors = ["#1F77B4", …], add /visualStyles/*/*/title = [{"fontSize": 14}]. Theme values are plain JSON (not expressions).',
      args: z.object({ operations: ops }),
      run: async ({ operations }) => patch(await themeFile(), operations as PatchOp[]),
    }),
    tool({
      name: "get_page",
      description: "Reads a page's page.json (size, background, filters, interactions).",
      args: z.object({ page: pageArg }),
      readOnly: true,
      run: async ({ page }) => read(`${P}/${(await pageOf(page)).name}/page.json`),
    }),
    tool({
      name: "create_page",
      description: "Adds a page (empty) at the end, or after another page.",
      args: z.object({
        display_name: z.string(),
        width: z.number().optional(),
        height: z.number().optional(),
        after: z.string().optional().describe("Page name or display name"),
        background: z.string().optional().describe('Page color behind the visuals, e.g. "#F4F6FA"'),
      }),
      run: async ({ display_name, width, height, after, background }) => {
        const f = await files();
        const name = newName();
        const first = (await pagesOf(f))[0]?.json;
        const json = buildPage(name, display_name, width ?? first?.width ?? 1280, height ?? first?.height ?? 720, first?.$schema);
        if (background) pageFormatOps(json, { background, wallpaper: background }).forEach((op) => applyOp(json, op));
        await putDraftFile(ref(), save, `${P}/${name}/page.json`, json);
        const order: string[] = [...((f.get(PAGES) as any)?.pageOrder ?? [])];
        order.splice(after ? order.indexOf((await pageOf(after)).name) + 1 : order.length, 0, name);
        await putDraftFile(ref(), save, PAGES, { ...(f.get(PAGES) as object), pageOrder: order });
        return `Page "${display_name}" created (name=${name}).`;
      },
    }),
    tool({
      name: "format_page",
      description: 'Sets a page\'s colors: background (behind the visuals) and wallpaper (around the page), e.g. {"background": "#F4F6FA"}.',
      args: z.object({ page: pageArg, background: z.string().optional(), wallpaper: z.string().optional() }),
      run: async ({ page, background, wallpaper }) => {
        const p = await pageOf(page);
        const file = `${P}/${p.name}/page.json`;
        await patch(file, pageFormatOps((await files()).get(file), { background, wallpaper }));
        return `Page "${p.json.displayName}" colors set.`;
      },
    }),
    tool({
      name: "update_page",
      description: "Changes a page's page.json with JSON Patch operations (displayName, size, background, filters, interactions).",
      args: z.object({ page: pageArg, operations: ops }),
      run: async ({ page, operations }) => patch(`${P}/${(await pageOf(page)).name}/page.json`, operations as PatchOp[]),
    }),
    tool({
      name: "delete_page",
      description: "Deletes a page and its visuals.",
      args: z.object({ page: z.string() }),
      run: async ({ page }) => {
        const p = await pageOf(page);
        const f = await files();
        const meta = f.get(PAGES) as any;
        if ((meta?.pageOrder ?? []).length <= 1) throw new Error("A report keeps at least one page.");
        await deleteDraftFiles(ref(), save, `${P}/${p.name}/`);
        const pageOrder = meta.pageOrder.filter((n: string) => n !== p.name);
        await putDraftFile(ref(), save, PAGES, { ...meta, pageOrder, activePageName: meta.activePageName === p.name ? pageOrder[0] : meta.activePageName });
        return `Page "${p.json.displayName}" deleted.`;
      },
    }),
    tool({
      name: "duplicate_page",
      description: "Copies a page with its visuals (new names), placed after it.",
      args: z.object({ page: z.string(), display_name: z.string() }),
      run: async ({ page, display_name }) => {
        const p = await pageOf(page);
        const f = await files();
        const name = newName();
        await putDraftFile(ref(), save, `${P}/${name}/page.json`, { ...structuredClone(p.json), name, displayName: display_name, visualInteractions: undefined });
        for (const v of visualsOf(f, p.name)) {
          const id = newName();
          await putDraftFile(ref(), save, `${P}/${name}/visuals/${id}/visual.json`, { ...structuredClone(v.json), name: id, parentGroupName: undefined });
        }
        const meta = f.get(PAGES) as any;
        const order = [...meta.pageOrder];
        order.splice(order.indexOf(p.name) + 1, 0, name);
        await putDraftFile(ref(), save, PAGES, { ...meta, pageOrder: order });
        return `Page "${display_name}" created as a copy (name=${name}).`;
      },
    }),
    tool({
      name: "reorder_pages",
      description: "Sets the order of the pages (every page, names or display names).",
      args: z.object({ pages: z.array(z.string()) }),
      run: async ({ pages }) => {
        const f = await files();
        const all = await pagesOf(f);
        const order = await Promise.all(pages.map(async (p) => (await pageOf(p)).name));
        if (order.length !== all.length || new Set(order).size !== all.length) throw new Error(`Give every page once: ${all.map((p) => p.json.displayName).join(", ")}.`);
        await putDraftFile(ref(), save, PAGES, { ...(f.get(PAGES) as object), pageOrder: order });
        return "Pages reordered.";
      },
    }),
    tool({
      name: "change_current_page",
      description: "Shows another page in the editor (a page not saved yet opens after your answer).",
      args: z.object({ page: z.string() }),
      run: async ({ page }) => {
        const p = await pageOf(page);
        const live = (await ctx.report()?.getPages())?.find((x) => x.name === p.name);
        if (live) {
          await live.setActive();
          return `Page "${p.json.displayName}" shown.`;
        }
        // A new page: made the report's opening page, so the editor shows it once reloaded.
        await patch(PAGES, [{ op: "add", path: "/activePageName", value: p.name }]);
        return `Page "${p.json.displayName}" opens after your answer (when the changes are saved).`;
      },
    }),
    tool({
      name: "get_visual",
      description: "Reads a visual's visual.json (fields, sort, formatting, filters).",
      args: z.object({ visual: visualArg, page: pageArg }),
      readOnly: true,
      run: async ({ visual, page }) => read((await visualOf(visual, page)).visual.file),
    }),
    tool({
      name: "create_visual",
      description: `Creates a visual from a spec (the JSON is written for you; field names are checked). Types and their roles: ${Object.entries(ROLES).map(([t, r]) => `${t} (${r.join(", ") || "no field"})`).join("; ")}. ${TYPE_NOTES}. Fields: 'Table'[Column], [Measure] or Sum('Table'[Column]) (sum, avg, count, min, max). "format": as in format_visual, applied at creation.`,
      args: z.object({
        page: pageArg,
        type: z.string(),
        x: z.number().optional().describe("Free position: only when layout_page cannot say it (leave x, y, width, height out, then call layout_page)"),
        y: z.number().optional(),
        width: z.number().optional(),
        height: z.number().optional(),
        fields: z.record(z.string(), z.array(z.string())).optional().describe('By role, e.g. {"Category": ["\'Sales\'[Region]"], "Y": ["[Revenue]"]}'),
        title: z.string().optional(),
        sort: z.object({ field: z.string(), descending: z.boolean().optional() }).optional(),
        text: z.string().optional().describe("textbox: its text"),
        format: formatArg.optional(),
        top: z.object({ count: z.number(), by: z.string().describe("Measure or Sum('T'[col]) ordering the items"), bottom: z.boolean().optional() }).optional().describe("Keep the first N items of the category (Top N), e.g. {\"count\": 10, \"by\": \"[Total Sales]\"}"),
      }),
      run: async ({ page, format, x, y, width, height, ...spec }) => {
        const p = await pageOf(page);
        const placed = x !== undefined && y !== undefined && width && height;
        const box = placed ? { x, y, width, height } : { ...freeSpot(visualsOf(await files(), p.name), p.json), ...(width && { width }), ...(height && { height }) };
        const { name, ignored } = await writeVisual(p.name, { ...spec, ...box }, format as Format | undefined);
        return `Visual created (name=${name}) on "${p.json.displayName}"${placed ? "" : " (placed for now: arrange the page with layout_page)"}.${ignoredText(ignored)}`;
      },
    }),
    tool({
      name: "layout_page",
      description: `${LAYOUT_HELP} margin / gap: px (default 20 / 16); top: space kept at the top (e.g. for a title bar).`,
      args: z.object({
        page: pageArg,
        rows: z.array(z.object({ height: z.number().optional(), items: z.array(z.object({ visual: z.string().describe("Visual name (id) or title"), width: z.number().optional().describe("Relative weight (default 1)") })) })),
        margin: z.number().optional(),
        gap: z.number().optional(),
        top: z.number().optional(),
      }),
      run: async ({ page, rows, margin, gap, top }) => {
        const p = await pageOf(page);
        const visuals = visualsOf(await files(), p.name);
        // Names first (a title is accepted when no other visual has it).
        const named = rows.map((row) => ({ ...row, items: row.items.map((i) => ({ ...i, visual: findVisual(visuals, i.visual).name })) }));
        const boxes = layoutRows(named, { width: p.json.width ?? 1280, height: p.json.height ?? 720 }, { margin, gap, top });
        for (const b of boxes) {
          const v = visuals.find((x) => x.name === b.visual)!;
          await patch(v.file, [{ op: "add", path: "/position", value: { ...v.json.position, x: b.x, y: b.y, width: b.width, height: b.height } }]);
        }
        const left = visuals.filter((v) => !v.json.isHidden && !v.json.parentGroupName && !boxes.some((b) => b.visual === v.name));
        return `Placed ${boxes.length} visual(s): ${boxes.map((b) => `${b.visual} x=${b.x} y=${b.y} w=${b.width} h=${b.height}`).join("; ")}.${left.length ? ` Not in the rows (kept where they are): ${left.map((v) => v.name).join(", ")}.` : ""}`;
      },
    }),
    tool({
      name: "create_button",
      description: 'Creates a button with an action: "page" (go to a page of the report), "back" (previous page) or "url" (open a web link). Leave x, y, width, height out to place it with layout_page.',
      args: z.object({
        page: pageArg,
        text: z.string(),
        action: z.object({ type: z.enum(["page", "back", "url"]), page: z.string().optional().describe("page: the target page (name or display name)"), url: z.string().optional(), tooltip: z.string().optional() }),
        fill: z.string().optional().describe('Button color "#RRGGBB"'),
        font_color: z.string().optional(),
        x: z.number().optional(),
        y: z.number().optional(),
        width: z.number().optional(),
        height: z.number().optional(),
      }),
      run: async ({ page, text, action, fill, font_color, x, y, width, height }) => {
        const p = await pageOf(page);
        const f = await files();
        const others = visualsOf(f, p.name);
        const target = action.type === "page" ? await pageOf(action.page) : undefined;
        const spot = x !== undefined && y !== undefined ? { x, y, width: width ?? 160, height: height ?? 40 } : { ...freeSpot(others, p.json), width: width ?? 160, height: height ?? 40 };
        const name = newName();
        const z = Math.max(0, ...others.map((v) => Number(v.json.position?.z ?? 0))) + 1000;
        const json = buttonVisual({ name, z, ...spot, text, fill, fontColor: font_color, action: action as Action, pageName: target?.name, schema: others[0]?.json.$schema ?? "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.14.0/schema.json" });
        await putDraftFile(ref(), save, `${P}/${p.name}/visuals/${name}/visual.json`, json);
        return `Button created (name=${name}) on "${p.json.displayName}"${target ? `, going to "${target.json.displayName}"` : ""}.`;
      },
    }),
    tool({
      name: "set_visual_action",
      description: 'Gives any visual (button, shape, card, image) a click action: "page", "back" or "url"; "none" removes it.',
      args: z.object({ visual: visualArg, page: pageArg, action: z.object({ type: z.enum(["page", "back", "url", "none"]), page: z.string().optional(), url: z.string().optional(), tooltip: z.string().optional() }) }),
      run: async ({ visual, page, action }) => {
        const { visual: v } = await visualOf(visual, page);
        if (action.type === "none") return patch(v.file, [{ op: "add", path: "/visual/visualContainerObjects/visualLink", value: [{ properties: { show: { expr: { Literal: { Value: "false" } } } } }] }]);
        const target = action.type === "page" ? await pageOf(action.page) : undefined;
        return patch(v.file, [{ op: "add", path: "/visual/visualContainerObjects", value: { ...(v.json.visual?.visualContainerObjects ?? {}), visualLink: linkProperties(action as Action, target?.name) } }]);
      },
    }),
    tool({
      name: "set_conditional_format",
      description: 'Conditional formatting of a visual: a table / matrix colors the cells of a field ("background" or "font"), a chart its bars / points, a card its value. One of: "gradient" {min, max, mid?} colors; "rules" [{op: ">", value: 0, color: "#2E7D32"}, …] (first matching rule wins); "from_field": true with based_on = a measure returning "#RRGGBB". based_on: the value deciding the color (default: field itself). Replaces the earlier format of that target.',
      args: z.object({
        visual: visualArg,
        page: pageArg,
        field: z.string().describe("A field of the visual, as in create_visual"),
        target: z.enum(["background", "font", "color"]).optional(),
        based_on: z.string().optional(),
        gradient: z.object({ min: z.string(), max: z.string(), mid: z.string().optional() }).optional(),
        rules: z.array(z.object({ op: z.enum([">", ">=", "<", "<=", "="]), value: z.number(), color: z.string() })).optional(),
        from_field: z.boolean().optional(),
      }),
      run: async ({ visual, page, ...spec }) => {
        const { visual: v } = await visualOf(visual, page);
        await patch(v.file, conditionalOps(v.json, spec, ctx.info(), measuresOf(await files())));
        return `Conditional format set on ${v.name} (${spec.field}).`;
      },
    }),
    tool({
      name: "set_reference_lines",
      description: 'Constant lines of a chart: a target, a threshold, an average… value = a number or a measure ([Target]). axis "y" (default) or "x". Replaces the lines of that axis; [] removes them.',
      args: z.object({
        visual: visualArg,
        page: pageArg,
        lines: z.array(z.object({ value: z.union([z.number(), z.string()]), axis: z.enum(["y", "x"]).optional(), label: z.string().optional(), color: z.string().optional(), style: z.enum(["solid", "dashed", "dotted"]).optional(), show_label: z.boolean().optional() })),
      }),
      run: async ({ visual, page, lines }) => {
        const { visual: v } = await visualOf(visual, page);
        const operations = lines.length ? referenceLineOps(v.json, lines, ctx.info(), measuresOf(await files())) : ["y1AxisReferenceLine", "xAxisReferenceLine"].filter((o) => v.json.visual?.objects?.[o]).map((o) => ({ op: "remove" as const, path: `/visual/objects/${o}` }));
        if (operations.length) await patch(v.file, operations);
        return `${lines.length} reference line(s) on ${v.name}.`;
      },
    }),
    tool({
      name: "copy_report_page",
      description: "Rebuilds a page of a reference report (Reports tab) in this report: same visual types, titles, fields and layout (scaled to the page), even when that report's definition is closed. Fields missing in this model are left out and listed. Formatting is not copied: format the new visuals after.",
      args: z.object({ report: z.string().describe("Reference report name"), page: z.string().describe("Its page (display name)"), display_name: z.string().optional().describe("Name of the new page (default: the same)") }),
      run: async ({ report, page, display_name }) => {
        const list = (await ctx.references?.()) ?? [];
        const source = list.find((x) => x.name.toLowerCase() === report.toLowerCase()) ?? list.find((x) => x.name.toLowerCase().includes(report.toLowerCase()));
        if (!source) throw new Error(list.length ? `No reference report "${report}". Reference reports: ${list.map((x) => `"${x.name}"`).join(", ")}.` : "No reference report: the user chooses them in the Reports tab.");
        if (!source.schema) throw new Error(`The structure of "${source.name}" could not be read (Reports screen › regenerate its schema).`);
        let f = await files();
        const first = (await pagesOf(f))[0]?.json;
        const size = { width: first?.width ?? 1280, height: first?.height ?? 720 };
        const copy = pageCopy(source.schema, page, ctx.info(), measuresOf(f), size);
        // The reference report's own measures the visuals use: made here when missing (checked by check_report).
        const missing = copy.measures.filter((m) => !measuresOf(f).some((x) => x.name === m.name));
        if (missing.length) {
          const ext = structuredClone((f.get(EXTENSION) as any) ?? { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/reportExtension/1.0.0/schema.json", name: "extension", entities: [] });
          for (const m of missing) {
            const entity = ext.entities.find((e: any) => e.name === m.table) ?? (ext.entities.push({ name: m.table, measures: [] }), ext.entities.at(-1));
            entity.measures.push({ name: m.name, dataType: "Double", expression: m.expression });
          }
          await putDraftFile(ref(), save, EXTENSION, ext);
          f = await files();
        }
        const name = newName();
        await putDraftFile(ref(), save, `${P}/${name}/page.json`, buildPage(name, display_name ?? copy.page.displayName, size.width, size.height, first?.$schema));
        await putDraftFile(ref(), save, PAGES, { ...(f.get(PAGES) as object), pageOrder: [...((f.get(PAGES) as any)?.pageOrder ?? []), name] });
        const made: string[] = [];
        const failed: string[] = [];
        for (const { source: id, ...spec } of copy.visuals) {
          try {
            made.push(`${(await writeVisual(name, spec)).name} (${spec.type}${spec.title ? ` "${spec.title}"` : ""})`);
          } catch (e) {
            failed.push(`${spec.type} ${id}: ${e instanceof Error ? e.message : e}`);
          }
        }
        return [
          `Page "${display_name ?? copy.page.displayName}" (name=${name}) rebuilt from "${source.name}": ${made.length} visual(s): ${made.join(", ") || "none"}.`,
          missing.length && `Report measures copied from it: ${missing.map((m) => `[${m.name}]`).join(", ")}.`,
          [...copy.dropped, ...failed].length && `Left out: ${[...copy.dropped, ...failed].join("; ")}.`,
          "Next: format the visuals (format_visual / update_theme) in the style asked, then change_current_page.",
        ]
          .filter(Boolean)
          .join("\n");
      },
    }),
    tool({
      name: "check_report",
      description: "Checks the draft by code: visuals overlapping or out of the page, report measures whose DAX fails. Also run for you before saving.",
      args: z.object({}),
      readOnly: true,
      run: async () => {
        const problems = await draftProblems(ctx, save);
        return problems.length ? `Problems:\n${problems.map((x) => `- ${x}`).join("\n")}` : "No problem found.";
      },
    }),
    tool({
      name: "set_visual_top_n",
      description: "Keeps only the first N items of a visual's category by a value (Top N filter of the visual); count 0 removes it.",
      args: z.object({ visual: visualArg, page: pageArg, count: z.number(), by: z.string().optional().describe("Measure or Sum('T'[col])"), bottom: z.boolean().optional() }),
      run: async ({ visual, page, count, by, bottom }) => {
        const { visual: v } = await visualOf(visual, page);
        const others = (v.json.filterConfig?.filters ?? []).filter((f: any) => f.type !== "TopN");
        if (count > 0 && !by) throw new Error('Give "by": the value ordering the items.');
        const filters = count > 0 ? [...others, topNFilter(v.json.visual?.query?.queryState ?? {}, { count, by: by!, bottom }, ctx.info(), measuresOf(await files()))] : others;
        return patch(v.file, [{ op: "add", path: "/filterConfig", value: { ...(v.json.filterConfig ?? {}), filters } }]);
      },
    }),
    tool({
      name: "set_visual_fields",
      description: "Replaces the fields of roles of a visual (other roles kept; an empty list removes a role), checked against the model.",
      args: z.object({ visual: visualArg, page: pageArg, fields: z.record(z.string(), z.array(z.string())) }),
      run: async ({ visual, page, fields }) => {
        const { visual: v } = await visualOf(visual, page);
        const type = v.json.visual?.visualType;
        const state = { ...(v.json.visual?.query?.queryState ?? {}) };
        const built = buildVisual({ type, x: 0, y: 0, width: 1, height: 1, fields: Object.fromEntries(Object.entries(fields).filter(([, l]) => l.length)) }, v.name, 0, ctx.info(), measuresOf(await files()));
        const fresh = ((built.visual as any).query?.queryState ?? {}) as Record<string, unknown>;
        for (const [role, list] of Object.entries(fields)) {
          if (list.length) state[roleKey(role)] = fresh[roleKey(role)];
          else delete state[roleKey(role)];
        }
        return patch(v.file, [{ op: "add", path: "/visual/query/queryState", value: state }]);
      },
    }),
    tool({
      name: "format_visual",
      description: 'Formats one or several visuals with plain values (no expressions to write): {"title": {"fontSize": 16, "bold": true, "fontColor": "#203247"}, "background": {"show": true, "color": "#FFFFFF"}, "border": {"show": true, "radius": 12}, "labels": {"show": true}}. Names are checked (get_format_options lists them); the same format goes to every visual given.',
      args: z.object({ visuals: z.array(z.string()).describe("Visual names (ids) or titles"), page: pageArg, format: formatArg }),
      run: async ({ visuals, page, format }) => {
        const targets = await Promise.all(visuals.map((v) => visualOf(v, page)));
        // Every visual checked before any is changed.
        const planned = targets.map(({ visual: v }) => ({ v, ...formatOps(v.json, format as Format) }));
        for (const { v, ops } of planned) if (ops.length) await patch(v.file, ops);
        return `Formatted ${planned.length} visual(s): ${planned.map(({ v, ops }) => `${v.name} (${ops.length} properties)`).join(", ")}.${ignoredText([...new Set(planned.flatMap((p) => p.ignored))])}`;
      },
    }),
    tool({
      name: "set_slicer_type",
      description: 'Changes how a slicer looks: "dropdown", "list" (vertical list), "tiles" (buttons side by side), "between" (a range, numbers / dates), "relative" (relative dates); single_select: one choice at a time.',
      args: z.object({ visual: visualArg, page: pageArg, type: z.enum(["dropdown", "list", "tiles", "between", "relative"]), single_select: z.boolean().optional() }),
      run: async ({ visual, page, type, single_select }) => {
        const { visual: v } = await visualOf(visual, page);
        if (v.json.visual?.visualType !== "slicer") throw new Error(`${v.name} is a ${v.json.visual?.visualType}, not a slicer.`);
        const mode = { dropdown: "Dropdown", list: "Basic", tiles: "Basic", between: "Between", relative: "Relative" }[type];
        const format: Format = { data: { mode }, ...(single_select !== undefined && { selection: { singleSelect: single_select } }) };
        const ops: PatchOp[] = formatOps(v.json, format).ops;
        // Orientation is the slicer's own "general" object (not the frame's), a number: written here.
        if (type === "list" || type === "tiles") {
          const general = v.json.visual?.objects?.general ?? [];
          const at = Math.max(0, general.findIndex((e: any) => !e?.selector));
          ops.push({ op: "add", path: `/visual/objects/general/${at}/properties/orientation`, value: { expr: { Literal: { Value: type === "tiles" ? "1D" : "0D" } } } });
        }
        await patch(v.file, ops);
        return `Slicer ${v.name}: ${type}${single_select ? ", single choice" : ""}.`;
      },
    }),
    tool({
      name: "update_visual",
      description: "Changes a visual's visual.json with JSON Patch operations (position, sort, filters, isHidden, per-series formatting with a selector). For plain formatting, format_visual is simpler.",
      args: z.object({ visual: visualArg, page: pageArg, operations: ops }),
      run: async ({ visual, page, operations }) => patch((await visualOf(visual, page)).visual.file, operations as PatchOp[]),
    }),
    tool({
      name: "delete_visual",
      description: "Deletes a visual.",
      args: z.object({ visual: visualArg, page: pageArg }),
      run: async ({ visual, page }) => {
        const { visual: v } = await visualOf(visual, page);
        await deleteDraftFiles(ref(), save, v.file.replace(/visual\.json$/, ""));
        return `Visual ${v.name} deleted.`;
      },
    }),
    tool({
      name: "duplicate_visual",
      description: "Copies a visual (same fields and formatting), at another place or on another page.",
      args: z.object({ visual: visualArg, page: pageArg, to_page: z.string().optional(), x: z.number().optional(), y: z.number().optional() }),
      run: async ({ visual, page, to_page, x, y }) => {
        const { visual: v } = await visualOf(visual, page);
        const target = await pageOf(to_page ?? page);
        const name = newName();
        const json = structuredClone(v.json);
        json.name = name;
        delete json.parentGroupName;
        json.position = { ...json.position, x: x ?? json.position.x + 20, y: y ?? json.position.y + 20, z: Number(json.position.z ?? 0) + 1 };
        await putDraftFile(ref(), save, `${P}/${target.name}/visuals/${name}/visual.json`, json);
        return `Visual copied (name=${name}) on "${target.json.displayName}".`;
      },
    }),
    tool({
      name: "get_visual_data",
      description: "The data a visual of the current page shows in the editor (first rows), as last saved: a visual created or changed in your draft shows its new data only after your answer (check its values with run_dax instead).",
      args: z.object({ visual: visualArg }),
      readOnly: true,
      run: async ({ visual }) => {
        const page = (await ctx.report()?.getPages())?.find((p) => p.isActive);
        if (!page) throw new Error("The report is not loaded.");
        const live = await page.getVisuals();
        const drafted = visualsOf(await files(), page.name).find((v) => v.name === visual || v.title?.toLowerCase() === visual.toLowerCase());
        // In the draft only: not in the editor yet (said plainly, so the AI does not search for it).
        if (drafted && !live.some((v) => v.name === drafted.name)) throw new Error(`Visual ${drafted.name} is in your draft, not saved yet: its data shows after your answer. Check its values with run_dax.`);
        const text = await visualText(findVisual(live, drafted?.name ?? visual), true, false, ctx.dataset());
        const changed = drafted && pendingFiles(ref()).some((f) => f.includes(`/visuals/${drafted.name}/`));
        return changed ? `${text}\n(As last saved: your draft changes this visual; its new data shows after your answer.)` : text;
      },
    }),
    tool({
      name: "get_format_options",
      description: 'The format objects of a visual type, or the properties of some of its objects ("title, labels") with their types and allowed values.',
      args: z.object({ visual_type: z.string(), object: z.string().optional() }),
      readOnly: true,
      run: async ({ visual_type, object }) => formatOptions(visual_type, object),
    }),
    tool({
      name: "update_measure",
      description: "Creates or changes a report measure (stored in the report, not in the model). Its DAX is evaluated on the model first: an error comes back as is.",
      args: z.object({ table: z.string().describe("Home table of the measure (a table of the model)"), name: z.string(), expression: z.string(), format: z.string().optional().describe('e.g. "#,0", "0.0%"') }),
      run: async ({ table, name, expression, format }) => {
        const f = await files();
        const ext = structuredClone((f.get(EXTENSION) as any) ?? { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/reportExtension/1.0.0/schema.json", name: "extension", entities: [] });
        const others = measuresOf(f).filter((m) => m.name !== name) as (ReportMeasures[number] & { expression: string })[];
        const dataset = ctx.dataset();
        let value: unknown;
        if (dataset) {
          const defs = [...others.map((m) => `MEASURE '${m.table}'[${m.name}] = ${m.expression}`), `MEASURE '${table}'[${name}] = ${expression}`];
          const [row] = await dax(dataset, `DEFINE\n${defs.join("\n")}\nEVALUATE ROW("value", [${name}])`);
          value = row?.value;
        }
        const dataType = value === null || value === undefined ? "Text" : typeof value === "boolean" ? "Boolean" : typeof value === "number" ? (Number.isInteger(value) ? "Integer" : "Double") : /^\d{4}-\d\d-\d\dT/.test(String(value)) ? "DateTime" : "Text";
        const measure = { name, dataType, expression, ...(format && { formatString: format }) };
        ext.entities = (ext.entities ?? []).map((e: any) => ({ ...e, measures: (e.measures ?? []).filter((m: any) => m.name !== name) })).filter((e: any) => e.measures.length || e.name === table);
        const entity = ext.entities.find((e: any) => e.name === table);
        if (entity) entity.measures.push(measure);
        else ext.entities.push({ name: table, measures: [measure] });
        await putDraftFile(ref(), save, EXTENSION, ext);
        return `Report measure '${table}'[${name}] written (value now: ${value ?? "blank"}).`;
      },
    }),
    tool({
      name: "delete_measure",
      description: "Deletes a report measure (visuals using it lose that field).",
      args: z.object({ name: z.string() }),
      run: async ({ name }) => {
        const ext = await readDraftJson(ref(), save, EXTENSION);
        if (!ext || !measuresOf(new Map([[EXTENSION, ext]])).some((m) => m.name === name)) throw new Error(`No report measure "${name}".`);
        ext.entities = ext.entities.map((e: any) => ({ ...e, measures: e.measures.filter((m: any) => m.name !== name) })).filter((e: any) => e.measures.length);
        await putDraftFile(ref(), save, EXTENSION, ext);
        return `Report measure [${name}] deleted.`;
      },
    }),
    tool({
      name: "discard_changes",
      description: "Drops your unsaved changes (the draft), e.g. after a wrong start.",
      args: z.object({}),
      run: async () => {
        const n = pendingFiles(ref()).length;
        discardDraft(ref());
        return n ? `Unsaved changes to ${n} file(s) dropped.` : "Nothing was pending.";
      },
    }),
    tool({
      name: "undo_last_save",
      description: "Puts the report back as it was before your last saved changes (only when the user asks to undo them).",
      args: z.object({}),
      run: async () => {
        const text = await undoLastSave(ref());
        await ctx.report()?.reload();
        return text;
      },
    }),
    queryTool(() => ctx.dataset() && { ...ctx.dataset()!, info: ctx.info() }, undefined, async () => measuresOf(await files()) as DefinedMeasure[]),
    tool({
      name: "run_dax",
      description: "Runs a DAX query (EVALUATE …) on the report's model, report measures included (max 50 rows shown): check names, values and totals before using them.",
      args: z.object({ query: z.string() }),
      readOnly: true,
      run: async ({ query }) => {
        const d = ctx.dataset();
        if (!d) throw new Error("The report's model is not readable.");
        // Report measures live in the report, not the model: all defined in the query (they may use each other).
        const own = (measuresOf(await files()) as (ReportMeasures[number] & { expression?: string })[]).filter((m) => m.expression);
        const defs = own.map((m) => `MEASURE '${m.table}'[${m.name}] = ${m.expression}`).join("\n");
        const q = !defs ? query : /^\s*DEFINE\b/i.test(query) ? query.replace(/^\s*DEFINE\b/i, `DEFINE\n${defs}`) : `DEFINE\n${defs}\n${query}`;
        return rowsText(await dax(d, q));
      },
    }),
  ];
}
