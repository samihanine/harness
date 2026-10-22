/**
 * The Builder when a report's definition cannot be read or written by the Fabric API (e.g. an
 * encrypted sensitivity label, or the older PBIR-Legacy format): the AI edits the report through the
 * Power BI editor embedded in the page (powerbi-report-authoring) — pages, visuals, fields, formatting,
 * Top N, layout — live, then the report is saved when its answer is done. No report measures, no theme
 * file in this mode (the definition stays closed).
 */
import { z } from "zod";
import { reportContext } from "./context";
import { tool, type Agent, type Tool } from "./loop";
import { modelTools } from "./model-tools";
import { DAX_RULES, daxTool, viewerTools } from "./pbi-tools";
import { resolveField } from "@/lib/model";
import { findVisual, models, pageText, type Report } from "@/lib/pbi";
import type { DatasetEntry, ReportEntry } from "@/lib/store";
import { ROLES, TYPE_NOTES } from "@/lib/visual-spec";

const PROPERTY = "http://powerbi.com/product/schema#property";
const AGGREGATIONS: Record<string, string> = { sum: "Sum", avg: "Average", average: "Average", count: "Count", countnonnull: "CountNonNull", min: "Min", max: "Max" };

/** 'Table'[Column], [Measure] or Sum('Table'[Column]) → a target of the authoring SDK, checked against the model. */
function target(spec: string, dataset?: DatasetEntry) {
  const m = spec.trim().match(/^(\w+)\((.+)\)$/);
  const known = resolveField(dataset?.info, m ? m[2] : spec);
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
}

/** The value of a format property as the SDK takes it. */
const value = (v: unknown) => ({ schema: PROPERTY, value: v });

const FORMAT_HELP =
  'By object, plain values: {"title": {"visible": true, "titleText": "Sales", "fontColor": "#203247", "textSize": 14, "alignment": "left"}, "background": {"visible": true, "color": "#FFFFFF", "transparency": 0}, "border": {"visible": true, "color": "#E5E7EB", "radius": 8}, "legend": {"visible": true, "position": "Top"}, "dataLabels": {"visible": true}, "categoryAxis": {"visible": true}, "valueAxis": {"visible": false}}';

export function builderSdkAgent({ entry, dataset, report, onSaved }: { entry: () => ReportEntry | undefined; dataset: () => DatasetEntry | undefined; report: () => Report | undefined; onSaved?: () => void }): Agent {
  const r = () => {
    const current = report();
    if (!current) throw new Error("The report is still loading in the editor: wait a moment and retry.");
    return current;
  };
  const datasetRef = () => entry() && { id: entry()!.datasetId, groupId: dataset()?.groupId ?? entry()!.datasetGroupId };
  let changed = false;
  const pageOf = async (page?: string) => {
    const pages = await r().getPages();
    const p = page ? pages.find((x) => x.name === page || x.displayName.toLowerCase() === page.toLowerCase()) : pages.find((x) => x.isActive);
    if (!p) throw new Error(`Page "${page}" not found. Pages: ${pages.map((x) => x.displayName).join(", ")}.`);
    return p;
  };
  const visualOf = async (visual: string, page?: string) => findVisual(await (await pageOf(page)).getVisuals(), visual);
  /** Every given format property, one by one: what the SDK refuses is listed (the rest is applied). */
  const format = async (v: Awaited<ReturnType<typeof visualOf>>, given: Record<string, Record<string, unknown>>) => {
    const refused: string[] = [];
    for (const [objectName, props] of Object.entries(given ?? {}))
      for (const [propertyName, v0] of Object.entries(props ?? {}))
        await v.setProperty({ objectName, propertyName }, value(v0)).catch((e: unknown) => refused.push(`${objectName}.${propertyName} (${e instanceof Error ? e.message : JSON.stringify(e)})`));
    return refused.length ? ` Not applied: ${refused.join("; ")}.` : "";
  };
  const setFields = async (v: Awaited<ReturnType<typeof visualOf>>, fields: Record<string, string[]>) => {
    // Every field checked before the visual changes.
    const targets = Object.entries(fields).map(([role, list]) => [role, list.map((f) => target(f, dataset()))] as const);
    for (const [role, list] of targets) {
      const current = await v.getDataFields(role).catch(() => []);
      for (let i = current.length - 1; i >= 0; i--) await v.removeDataField(role, i);
      for (const t of list) await v.addDataField(role, t as never);
    }
  };
  const topN = async (v: Awaited<ReturnType<typeof visualOf>>, top: { count: number; by: string; bottom?: boolean }) => {
    const roles = ["Category", "Rows", "Values", "Series"];
    let item: { table: string; column: string } | undefined;
    for (const role of roles) {
      const fields = (await v.getDataFields(role).catch(() => [])) as { table?: string; column?: string; aggregationFunction?: string }[];
      const f = fields.find((x) => x.column && !x.aggregationFunction);
      if (f) {
        item = { table: f.table!, column: f.column! };
        break;
      }
    }
    if (!item) throw new Error("A Top N needs a column in the visual (Category or Rows).");
    const m = await models();
    await v.setFilters([{ $schema: "http://powerbi.com/product/schema#topN", filterType: m.FilterType.TopN, target: item, operator: top.bottom ? "Bottom" : "Top", itemCount: Math.round(top.count), orderBy: target(top.by, dataset()) } as never]);
  };
  const fieldsArg = z.record(z.string(), z.array(z.string())).describe('By role, e.g. {"Category": ["\'Sales\'[Region]"], "Y": ["[Revenue]"]}');
  const formatArg = z.record(z.string(), z.record(z.string(), z.unknown())).describe(FORMAT_HELP);
  const topArg = z.object({ count: z.number(), by: z.string(), bottom: z.boolean().optional() });
  const mark = (t: Tool): Tool => (t.readOnly ? t : { ...t, run: async (args) => ((changed = true), t.run(args)) });
  const viewer = viewerTools(report, datasetRef, () => dataset()?.info).filter((t) => ["list_pages", "read_page", "change_current_page", "get_visual_data"].includes(t.name));

  const tools: Tool[] = [
    ...viewer,
    tool({
      name: "create_page",
      description: "Adds a page and shows it.",
      args: z.object({ display_name: z.string() }),
      run: async ({ display_name }) => {
        const p = await r().addPage(display_name);
        await p.setActive();
        return `Page "${display_name}" created and shown (name=${p.name}).`;
      },
    }),
    tool({
      name: "rename_page",
      description: "Renames a page.",
      args: z.object({ page: z.string(), display_name: z.string() }),
      run: async ({ page, display_name }) => {
        await (await pageOf(page)).setDisplayName(display_name);
        return "Page renamed.";
      },
    }),
    tool({
      name: "delete_page",
      description: "Deletes a page and its visuals.",
      args: z.object({ page: z.string() }),
      run: async ({ page }) => {
        const p = await pageOf(page);
        await r().deletePage(p.name);
        return `Page "${p.displayName}" deleted.`;
      },
    }),
    tool({
      name: "create_visual",
      description: `Creates a visual on a page (default: the current one) with its fields (checked against the model), title, format and Top N. Types and roles: ${Object.entries(ROLES).filter(([t]) => !["textbox", "shape"].includes(t)).map(([t, roles]) => `${t} (${roles.join(", ")})`).join("; ")}. ${TYPE_NOTES}. Fields: 'Table'[Column], [Measure] or Sum('Table'[Column]).`,
      args: z.object({ page: z.string().optional(), type: z.string(), x: z.number(), y: z.number(), width: z.number(), height: z.number(), fields: fieldsArg.optional(), title: z.string().optional(), format: formatArg.optional(), top: topArg.optional() }),
      run: async ({ page, type, x, y, width, height, fields, title, format: given, top }) => {
        const p = await pageOf(page);
        // Fields checked before anything is created (nothing half made on a wrong name).
        Object.values(fields ?? {}).flat().forEach((f) => target(f, dataset()));
        // Without an explicit display state the SDK makes the visual hidden.
        const { visual } = await p.createVisual(type, { x, y, width, height, displayState: { mode: 0 } });
        await p.setVisualDisplayState(visual.name, 0).catch(() => undefined);
        const v = findVisual(await p.getVisuals(), visual.name);
        if (fields) await setFields(v, fields);
        // A titled card would say its field twice (its label is also cut in a short box).
        const style = { ...(title && type === "card" && { categoryLabels: { visible: false } }), ...(given ?? {}), ...(title && { title: { visible: true, titleText: title, ...(given?.title ?? {}) } }) };
        const note = await format(v, style);
        if (top) await topN(v, top);
        return `Visual created (name=${visual.name}) on "${p.displayName}".${note}`;
      },
    }),
    tool({
      name: "set_visual_fields",
      description: "Replaces the fields of roles of a visual (others kept; an empty list empties the role).",
      args: z.object({ visual: z.string(), page: z.string().optional(), fields: fieldsArg }),
      run: async ({ visual, page, fields }) => {
        await setFields(await visualOf(visual, page), fields);
        return "Fields set.";
      },
    }),
    tool({
      name: "format_visual",
      description: `Formats one or several visuals (same format to each). ${FORMAT_HELP}`,
      args: z.object({ visuals: z.array(z.string()), page: z.string().optional(), format: formatArg }),
      run: async ({ visuals, page, format: given }) => {
        const notes: string[] = [];
        for (const name of visuals) notes.push(await format(await visualOf(name, page), given));
        return `Formatted ${visuals.length} visual(s).${[...new Set(notes)].join("")}`;
      },
    }),
    tool({
      name: "move_visual",
      description: "Moves and/or resizes a visual (page px).",
      args: z.object({ visual: z.string(), page: z.string().optional(), x: z.number().optional(), y: z.number().optional(), width: z.number().optional(), height: z.number().optional() }),
      run: async ({ visual, page, x, y, width, height }) => {
        const v = await visualOf(visual, page);
        if (x !== undefined || y !== undefined) await v.moveVisual(x ?? v.layout.x ?? 0, y ?? v.layout.y ?? 0);
        if (width !== undefined || height !== undefined) await v.resizeVisual(width ?? v.layout.width ?? 300, height ?? v.layout.height ?? 200);
        return "Visual moved.";
      },
    }),
    tool({
      name: "set_visual_top_n",
      description: "Keeps only the first (or last) N items of a visual's category by a value.",
      args: z.object({ visual: z.string(), page: z.string().optional(), count: z.number(), by: z.string(), bottom: z.boolean().optional() }),
      run: async ({ visual, page, ...top }) => {
        await topN(await visualOf(visual, page), top);
        return "Top N set.";
      },
    }),
    tool({
      name: "delete_visual",
      description: "Deletes a visual.",
      args: z.object({ visual: z.string(), page: z.string().optional() }),
      run: async ({ visual, page }) => {
        const p = await pageOf(page);
        const v = findVisual(await p.getVisuals(), visual);
        await p.deleteVisual(v.name);
        return `Visual ${v.name} deleted.`;
      },
    }),
    daxTool(datasetRef),
    ...modelTools(dataset),
  ].map(mark);

  return {
    instructions: `Screen: the BUILDER, in "Power BI editor" mode: this report's definition cannot be read by the API (e.g. a sensitivity label), so you change it through the editor shown to the user — pages, visuals, fields, formatting, Top N, layout. Changes show at once and the report is saved when your answer is done. Not possible here: report measures, the theme file (say so if asked; measures can be added to the model in the Datasets screen when the user may edit it).
How to work (few, well-filled calls; several independent calls in one reply):
- The page state lists the pages and the visuals of the current page; list_pages / read_page for the others.
- New page: create_page (it is shown), then create_visual for each visual with its title, fields and format, in one reply.
- KPI cards: one card per value. Top N: create_visual "top": {"count": 10, "by": "[Measure]"}.
- Layout on 1280×720: 20 px margins, 16 px gaps, a row of cards (y 20, height 110), charts below, inside the page, no overlap.
- Titles and texts in the user's language. Use only fields that exist (model summary).
${DAX_RULES}`,
    tools,
    context: () => reportContext(entry()!, dataset()),
    state: async () => (report() ? `Editor mode (definition closed). What the editor shows:\n${await pageText(report()!, false, datasetRef())}` : "The report is loading in the editor."),
    finish: async () => {
      if (!changed) return undefined;
      changed = false;
      await r().save();
      onSaved?.();
      return "✓ Report saved.";
    },
  };
}
