/**
 * The Builder in editor mode (reports whose definition the API cannot write: PBIR-Legacy format, or
 * refused): the report is changed live in the Power BI editor embedded on the user's screen (browser
 * actions, src/tools/browser/editor-actions.ts), then saved. Fields are checked against the model here.
 */
import { z } from "zod";
import { resolveField } from "@server/pbi/model";
import { modelInfo } from "@server/pbi/models";
import { layoutProblems, layoutRows, snapBox } from "@server/pbi/layout";
import { forgetReport } from "@server/capture/browser";
import { ROLES, TYPE_NOTES, fitRoles } from "@server/pbi/visual-spec";
import { reportOf } from "../helpers";
import { json, tool, type ToolContext } from "../tool";

const AGGREGATIONS: Record<string, string> = { sum: "Sum", avg: "Average", average: "Average", count: "Count", countnonnull: "CountNonNull", min: "Min", max: "Max" };

/** 'Table'[Column], [Measure] or Sum('Table'[Column]) → a target of the editor, checked against the model. */
async function targets(ctx: ToolContext, fields?: Record<string, string[]>, type?: string) {
  if (!fields) return undefined;
  const r = reportOf(ctx);
  const info = r.datasetId ? await modelInfo({ id: r.datasetId, groupId: r.datasetGroupId }).catch(() => undefined) : undefined;
  const one = (spec: string) => {
    const m = spec.trim().match(/^(\w+)\((.+)\)$/);
    const known = resolveField(info, m ? m[2] : spec);
    if (known.kind === "measure") {
      if (m) throw new Error(`[${known.name}] is a measure: use it without ${m[1]}().`);
      return { table: known.table, measure: known.name };
    }
    if (m) {
      const fn = AGGREGATIONS[m[1].toLowerCase()];
      if (!fn) throw new Error(`"${m[1]}" is not an aggregation (${Object.keys(AGGREGATIONS).join(", ")}).`);
      return { table: known.table, column: known.name, aggregationFunction: fn };
    }
    return { table: known.table, column: known.name };
  };
  const fitted = type ? (fitRoles(type, fields) ?? fields) : fields;
  return Object.fromEntries(Object.entries(fitted).map(([role, list]) => [role, list.map(one)]));
}

const fieldsArg = z.record(z.string(), z.array(z.string())).describe('By role, e.g. {"Category": ["\'Sales\'[Region]"], "Y": ["[Revenue]"]}');
const FORMAT_HELP =
  'Frame and general properties of the visual, plain values: {"title": {"visible": true, "titleText": "Sales", "fontColor": "#203247", "textSize": 14, "alignment": "left"}, "background": {"visible": true, "color": "#FFFFFF"}, "border": {"visible": true, "color": "#E5E7EB", "radius": 8}, "legend": {"visible": true, "position": "Top"}, "dataLabels": {"visible": true}}. The inside of charts (series colors, labels, table headers): update_theme "styles".';

export const editorOverview = tool({
  name: "report_overview",
  description: "The report as the editor shows it now (editor mode): pages (name, size), visuals (type, title, place, fields by role).",
  args: z.object({ page: z.string().optional() }),
  async run(a, ctx) {
    return `Editor mode (the report is changed live in the Power BI editor on screen; save_report saves it).\n${json(await ctx.browser("editor.overview", { page: a.page }))}`;
  },
});

export const editorAddPage = tool({
  name: "add_page",
  description: "Add a page to the report (editor mode) and show it. Returns its name.",
  args: z.object({ display_name: z.string() }),
  changes: true,
  async run(a, ctx) {
    const r = (await ctx.browser("editor.addPage", { displayName: a.display_name })) as { name: string };
    return `Page "${a.display_name}" added: name=${r.name}.`;
  },
});

export const editorAddVisual = tool({
  name: "add_visual",
  description: `Add a visual (editor mode) with its fields (checked against the model), title and format. Types and roles: ${Object.entries(ROLES).filter(([t]) => !["textbox", "shape", "actionButton"].includes(t)).map(([t, r]) => `${t} (${r.join(", ")})`).join("; ")}. ${TYPE_NOTES}. Then place the page's visuals with layout_page.`,
  args: z.object({ page: z.string().optional(), type: z.string(), fields: fieldsArg.optional(), title: z.string().optional(), format: z.record(z.string(), z.record(z.string(), z.unknown())).optional().describe(FORMAT_HELP), x: z.number().optional(), y: z.number().optional(), width: z.number().optional(), height: z.number().optional() }),
  changes: true,
  async run(a, ctx) {
    const fields = await targets(ctx, a.fields, a.type);
    const r = (await ctx.browser("editor.addVisual", { page: a.page, type: a.type, layout: { x: a.x ?? 20, y: a.y ?? 20, width: a.width ?? 400, height: a.height ?? 250 }, fields, title: a.title, format: a.format })) as { name: string; refused: string[] };
    return `Visual ${a.type} added: name=${r.name}.${r.refused.length ? ` Not applied: ${r.refused.join(", ")}.` : ""}`;
  },
});

export const editorUpdateVisual = tool({
  name: "update_visual",
  description: 'Change a visual in place (editor mode): "type" turns it into another type (fields kept), "fields" replaces the fields of the roles given.',
  args: z.object({ page: z.string().optional(), visual: z.string(), type: z.string().optional(), fields: fieldsArg.optional() }),
  changes: true,
  async run(a, ctx) {
    const r = await ctx.browser("editor.updateVisual", { page: a.page, visual: a.visual, type: a.type, fields: await targets(ctx, a.fields, a.type) });
    return `Visual now: ${json(r)}`;
  },
});

export const editorFormatVisual = tool({
  name: "format_visual",
  description: `Format visuals (editor mode), the same format to each. ${FORMAT_HELP}`,
  args: z.object({ page: z.string().optional(), visuals: z.array(z.string()), format: z.record(z.string(), z.record(z.string(), z.unknown())) }),
  changes: true,
  async run(a, ctx) {
    const r = (await ctx.browser("editor.format", { page: a.page, visuals: a.visuals, format: a.format })) as { refused: string[] };
    return `Formatted ${a.visuals.length} visual(s).${r.refused.length ? ` Not applied (use update_theme "styles" for these): ${r.refused.join("; ")}.` : ""}`;
  },
});

export const editorLayoutPage = tool({
  name: "layout_page",
  description: "Lay out a page in rows (editor mode), top to bottom: each row a height in px (omit to share the rest) and its visuals left to right with relative widths. Every visual of the page in a row.",
  args: z.object({ page: z.string().optional(), rows: z.array(z.object({ height: z.number().optional(), items: z.array(z.object({ visual: z.string(), width: z.number().optional() })) })), margin: z.number().optional(), gap: z.number().optional() }),
  changes: true,
  async run(a, ctx) {
    const [p] = (await ctx.browser("editor.overview", { page: a.page })) as { name: string; size?: { width: number; height: number }; visuals: { name: string; title?: string }[] }[];
    const named = a.rows.map((row) => ({ ...row, items: row.items.map((i) => ({ ...i, visual: p.visuals.find((v) => v.name === i.visual || (v.title ?? "").toLowerCase() === i.visual.toLowerCase())?.name ?? i.visual })) }));
    const boxes = layoutRows(named, { width: p.size?.width ?? 1280, height: p.size?.height ?? 720 }, { margin: a.margin ?? 20, gap: a.gap ?? 16 });
    await ctx.browser("editor.move", { page: p.name, boxes });
    return `Placed ${boxes.length} visuals.`;
  },
});

export const editorPlaceVisuals = tool({
  name: "place_visuals",
  description: "Free positions in px (editor mode): each visual's x, y, width, height, snapped to a 4 px grid and kept in the page; overlaps are reported.",
  args: z.object({ page: z.string().optional(), visuals: z.array(z.object({ visual: z.string(), x: z.number(), y: z.number(), width: z.number(), height: z.number() })) }),
  changes: true,
  async run(a, ctx) {
    type Page = { name: string; size?: { width: number; height: number }; visuals: { name: string; title?: string; type: string; x: number; y: number; width: number; height: number; hidden?: boolean }[] };
    const [p] = (await ctx.browser("editor.overview", { page: a.page })) as Page[];
    const size = { width: p.size?.width ?? 1280, height: p.size?.height ?? 720 };
    const boxes = a.visuals.map((b) => snapBox({ ...b, visual: p.visuals.find((v) => v.name === b.visual || (v.title ?? "").toLowerCase() === b.visual.toLowerCase())?.name ?? b.visual }, size));
    await ctx.browser("editor.move", { page: p.name, boxes });
    const placed = p.visuals.map((v) => ({ ...v, ...(boxes.find((b) => b.visual === v.name) ?? {}) }));
    const problems = layoutProblems(placed, size);
    return `Placed ${boxes.length} visual(s).${problems.length ? ` Overlaps: ${problems.join("; ")}.` : ""}`;
  },
});

export const editorRemove = tool({
  name: "remove",
  description: "Remove a visual (page + visual) or a whole page (page only), editor mode.",
  args: z.object({ page: z.string(), visual: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    await ctx.browser(a.visual ? "editor.deleteVisual" : "editor.deletePage", { page: a.page, visual: a.visual });
    return a.visual ? "Visual removed." : "Page removed.";
  },
});

export const editorRenamePage = tool({
  name: "rename_page",
  description: "Rename a page (editor mode).",
  args: z.object({ page: z.string(), display_name: z.string() }),
  changes: true,
  async run(a, ctx) {
    await ctx.browser("editor.renamePage", { page: a.page, displayName: a.display_name });
    return "Page renamed.";
  },
});

export const editorTheme = tool({
  name: "update_theme",
  description:
    'Change the report theme (editor mode): palette and "styles" — the inside of a visual type, for every visual of that type ("*": all). E.g. {"dataColors": ["#B39251", "#1C4358"], "styles": {"tableEx": {"columnHeaders": {"fontColor": "#FFFFFF", "backColor": "#363D4E"}}, "*": {"title": {"fontSize": 14, "bold": true}}}}.',
  args: z.object({ dataColors: z.array(z.string()).optional(), background: z.string().optional(), foreground: z.string().optional(), styles: z.record(z.string(), z.record(z.string(), z.record(z.string(), z.unknown()))).optional() }),
  changes: true,
  async run(a, ctx) {
    const value = (v: unknown) => (typeof v === "string" && /^#[0-9a-f]{3,8}$/i.test(v) ? { solid: { color: v } } : v);
    const { styles, ...colors } = a;
    const visualStyles = styles && Object.fromEntries(Object.entries(styles).map(([type, objects]) => [type, { "*": Object.fromEntries(Object.entries(objects).map(([o, props]) => [o, [Object.fromEntries(Object.entries(props).map(([p, v]) => [p, value(v)]))]])) }]));
    await ctx.browser("editor.theme", { theme: { ...Object.fromEntries(Object.entries(colors).filter(([, v]) => v !== undefined)), ...(visualStyles && { visualStyles }) } });
    return "Theme updated.";
  },
});

export const editorSave = tool({
  name: "save_report",
  description: "Save the report as the editor shows it (editor mode). Then look_at_report to check.",
  args: z.object({}),
  changes: true,
  async run(_, ctx) {
    ctx.status("Saving the report…");
    await ctx.browser("editor.save");
    // The headless browser's copy is the old version: the next look loads the saved one.
    await forgetReport(reportOf(ctx).id);
    await new Promise((r) => setTimeout(r, 2000));
    return "Saved.";
  },
});

export const EDITOR_TOOLS = [editorOverview, editorAddPage, editorAddVisual, editorUpdateVisual, editorFormatVisual, editorLayoutPage, editorPlaceVisuals, editorRemove, editorRenamePage, editorTheme, editorSave];
