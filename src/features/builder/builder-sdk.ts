/**
 * The Builder when a report's definition cannot be read or written by the Fabric API (e.g. an
 * encrypted sensitivity label, or the older PBIR-Legacy format): the AI edits the report through the
 * Power BI editor embedded in the page (powerbi-report-authoring) — pages, visuals, fields, formatting,
 * Top N, layout — live, then the report is saved when its answer is done. No report measures, no theme
 * file in this mode (the definition stays closed).
 */
import { z } from "zod";
import { reportContext } from "@/agent/context";
import { examplesText, referenceReports } from "@/services/report-schema";
import { LAYOUT_HELP, layoutRows } from "@/features/builder/layout";
import { pageCopy } from "@/features/builder/report-copy";
import { describeError, tool, type Agent, type Tool } from "@/agent/loop";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES, daxTool, queryTool, viewerTools } from "@/agent/pbi-tools";
import { resolveField } from "@/services/model";
import { findVisual, models, pageText, type Report } from "@/services/pbi";
import type { DatasetEntry, ReportEntry } from "@/lib/store";
import { ROLES, TYPE_NOTES, roleKey } from "@/features/builder/visual-spec";
import { BUILDER_EDITOR_INSTRUCTIONS } from "@/prompts/builder";

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
  'The frame of a visual, plain values: {"title": {"visible": true, "titleText": "Sales", "fontColor": "#203247", "textSize": 14, "alignment": "left"}, "background": {"visible": true, "color": "#FFFFFF", "transparency": 0}, "border": {"visible": true, "color": "#E5E7EB", "radius": 8}, "dropShadow": {"visible": true}, "visualHeader": {"visible": false}, "legend": {"visible": true, "position": "Top"}, "categoryAxis": {"visible": true}, "valueAxis": {"visible": false}}. The inside of a chart or matrix (headers, grid, values, labels, colors) is not reachable here: use update_theme "styles".';

/** A theme style value: "#RRGGBB" becomes the theme's color form, the rest is kept. */
const themeValue = (v: unknown) => (typeof v === "string" && /^#[0-9a-f]{3,8}$/i.test(v) ? { solid: { color: v } } : v);

/** Deep merge of plain objects (lists replaced). */
const merge = (a: any, b: any): any => {
  if (!a || typeof a !== "object" || Array.isArray(a) || !b || typeof b !== "object" || Array.isArray(b)) return b;
  const out = { ...a };
  for (const [k, v] of Object.entries(b)) out[k] = merge(a[k], v);
  return out;
};

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
        await v.setProperty({ objectName, propertyName }, value(v0)).catch((e: unknown) => refused.push(`${objectName}.${propertyName}${/unsupported/i.test(JSON.stringify(e)) ? "" : ` (${e instanceof Error ? e.message : JSON.stringify(e)})`}`));
    return refused.length ? ` Not applied (not reachable on a visual in this mode): ${refused.join(", ")} — for these, use update_theme "styles" (by visual type).` : "";
  };
  const setFields = async (v: Awaited<ReturnType<typeof visualOf>>, fields: Record<string, string[]>) => {
    // Every field checked before the visual changes.
    const targets = Object.entries(fields).map(([role, list]) => [role, list.map((f) => target(f, dataset()))] as const);
    for (const [given, list] of targets) {
      const role = roleKey(given);
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
  type Spec = { type: string; x?: number; y?: number; width?: number; height?: number; fields?: Record<string, string[]>; title?: string; format?: Record<string, Record<string, unknown>>; top?: { count: number; by: string; bottom?: boolean } };
  /** One visual on a page: fields checked first (nothing half made on a wrong name); without a place, under the others. */
  const createOne = async (p: Awaited<ReturnType<typeof pageOf>>, { type, x, y, width, height, fields, title, format: given, top }: Spec) => {
    Object.values(fields ?? {}).flat().forEach((f) => target(f, dataset()));
    let box = { x: x ?? 20, y: y ?? 20, width: width ?? 420, height: height ?? 240 };
    if (x === undefined || y === undefined) {
      const bottom = Math.max(0, ...(await p.getVisuals()).map((o) => (o.layout.y ?? 0) + (o.layout.height ?? 0)));
      box = { ...box, x: 20, y: Math.min(bottom + 16, Math.max(0, (p.defaultSize?.height ?? 720) - box.height)) };
    }
    // Without an explicit display state the SDK makes the visual hidden.
    const { visual } = await p.createVisual(type, { ...box, displayState: { mode: 0 } });
    await p.setVisualDisplayState(visual.name, 0).catch(() => undefined);
    const v = findVisual(await p.getVisuals(), visual.name);
    if (fields) await setFields(v, fields);
    // A titled card would say its field twice (its label is also cut in a short box).
    const style = { ...(title && type === "card" && { categoryLabels: { visible: false } }), ...(given ?? {}), ...(title && { title: { visible: true, titleText: title, ...(given?.title ?? {}) } }) };
    const note = await format(v, style);
    if (top) await topN(v, top);
    return { name: visual.name, note };
  };
  const mark = (t: Tool): Tool => (t.readOnly ? t : { ...t, run: async (args) => ((changed = true), t.run(args)) });
  // The viewer's reading and filter tools: in the editor, filters set this way are saved with the report.
  const viewer = viewerTools(report, datasetRef, () => dataset()?.info).filter((t) =>
    ["list_pages", "read_page", "change_current_page", "get_visual_data", "update_report_filter", "update_page_filter", "update_visual_filter", "update_slicer_value", "reset_filters"].includes(t.name),
  );

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
      name: "resize_page",
      description: "Changes a page's size in px (e.g. height 1400 for a long page read by scrolling; 1280×720 is the usual 16:9). Then arrange it again with layout_page.",
      args: z.object({ page: z.string().optional(), width: z.number(), height: z.number() }),
      run: async ({ page, width, height }) => {
        const p = await pageOf(page);
        if (width < 320 || height < 240 || width > 4000 || height > 8000) throw new Error("A page is between 320×240 and 4000×8000 px.");
        await p.resizePage((await models()).PageSizeType.Custom, Math.round(width), Math.round(height));
        return `Page "${p.displayName}" is now ${Math.round(width)}×${Math.round(height)}.`;
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
      args: z.object({ page: z.string().optional(), type: z.string(), x: z.number().optional().describe("Free position: only when layout_page cannot say it (leave x, y, width, height out, then layout_page)"), y: z.number().optional(), width: z.number().optional(), height: z.number().optional(), fields: fieldsArg.optional(), title: z.string().optional(), format: formatArg.optional(), top: topArg.optional() }),
      run: async ({ page, ...spec }) => {
        const p = await pageOf(page);
        const { name, note } = await createOne(p, spec);
        return `Visual created (name=${name}) on "${p.displayName}"${spec.x === undefined ? " (placed for now: arrange the page with layout_page)" : ""}.${note}`;
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
      name: "layout_page",
      description: `${LAYOUT_HELP} margin / gap: px (default 20 / 16); top: space kept at the top.`,
      args: z.object({
        page: z.string().optional(),
        rows: z.array(z.object({ height: z.number().optional(), items: z.array(z.object({ visual: z.string(), width: z.number().optional() })) })),
        margin: z.number().optional(),
        gap: z.number().optional(),
        top: z.number().optional(),
      }),
      run: async ({ page, rows, margin, gap, top }) => {
        const p = await pageOf(page);
        const visuals = await p.getVisuals();
        const named = rows.map((row) => ({ ...row, items: row.items.map((i) => ({ ...i, visual: findVisual(visuals, i.visual).name })) }));
        const boxes = layoutRows(named, { width: p.defaultSize?.width ?? 1280, height: p.defaultSize?.height ?? 720 }, { margin, gap, top });
        for (const b of boxes) {
          const v = visuals.find((x) => x.name === b.visual)!;
          await v.moveVisual(b.x, b.y);
          await v.resizeVisual(b.width, b.height);
        }
        return `Placed ${boxes.length} visual(s): ${boxes.map((b) => `${b.visual} x=${b.x} y=${b.y} w=${b.width} h=${b.height}`).join("; ")}.`;
      },
    }),
    tool({
      name: "copy_report_page",
      description: "Rebuilds a page of a reference report (Reports tab) in this report: same visual types, titles, fields and layout (scaled), even when that report's definition is closed. Fields missing in this model, and that report's own measures, are left out and listed. Formatting is not copied: format after.",
      args: z.object({ report: z.string(), page: z.string(), display_name: z.string().optional() }),
      run: async ({ report: name, page, display_name }) => {
        const list = await referenceReports(entry()?.exampleIds);
        const source = list.find((x) => x.name.toLowerCase() === name.toLowerCase()) ?? list.find((x) => x.name.toLowerCase().includes(name.toLowerCase()));
        if (!source) throw new Error(list.length ? `No reference report "${name}". Reference reports: ${list.map((x) => `"${x.name}"`).join(", ")}.` : "No reference report: the user chooses them in the Reports tab.");
        if (!source.schema) throw new Error(`The structure of "${source.name}" could not be read.`);
        const first = (await r().getPages())[0];
        // No report measures in this mode: the reference report's own measures are left out.
        const copy = pageCopy(source.schema, page, dataset()?.info, [], { width: first?.defaultSize?.width ?? 1280, height: first?.defaultSize?.height ?? 720 });
        const p = await r().addPage(display_name ?? copy.page.displayName);
        await p.setActive();
        const made: string[] = [];
        const failed: string[] = [];
        for (const { source: id, ...spec } of copy.visuals) {
          await createOne(p, spec).then(
            (v) => made.push(`${v.name} (${spec.type})`),
            (e) => failed.push(`${spec.type} ${id}: ${describeError(e)}`),
          );
        }
        return [`Page "${p.displayName}" rebuilt from "${source.name}": ${made.length} visual(s): ${made.join(", ") || "none"}.`, [...copy.dropped, ...failed].length && `Left out: ${[...copy.dropped, ...failed].join("; ")}.`].filter(Boolean).join("\n");
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
      name: "set_slicer_type",
      description:
        'Changes how a slicer looks: "tiles" (buttons), "list" (list slicer), "text" (search box) or "classic" — for that slicer only; "dropdown": classic slicers become dropdowns through the theme (EVERY classic slicer of the report: say so to the user).',
      args: z.object({ visual: z.string(), page: z.string().optional(), type: z.enum(["dropdown", "tiles", "list", "text", "classic"]) }),
      run: async ({ visual, page, type }) => {
        const v = await visualOf(visual, page);
        if (!/slicer/i.test(v.type)) throw new Error(`${v.name} is a ${v.type}, not a slicer.`);
        // The SDK cannot reach a slicer's own settings: its type changes, or the theme gives the dropdown.
        const kind = { tiles: "advancedSlicerVisual", list: "listSlicer", text: "textSlicer", classic: "slicer", dropdown: "slicer" }[type];
        if (v.type !== kind) await v.changeType(kind);
        if (type !== "dropdown") return `Slicer ${v.name}: ${type}.`;
        const current = ((await r().getTheme().catch(() => undefined)) as { themeJson?: Record<string, unknown> } | undefined)?.themeJson ?? { name: "Hub theme" };
        await r().applyTheme({ themeJson: merge(current, { visualStyles: { slicer: { "*": { data: [{ mode: "Dropdown" }] } } } }) });
        return `Slicer ${v.name}: dropdown (through the theme: every classic slicer of the report is now a dropdown).`;
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
    tool({
      name: "update_theme",
      description:
        'Changes the report theme (kept with the report when saved): palette, page colors, and "styles" — the formatting of the inside of a visual type, for every visual of that type ("*": all types). E.g. {"dataColors": ["#16A34A", "#2563EB"], "styles": {"pivotTable": {"columnHeaders": {"fontColor": "#FFFFFF", "backColor": "#16A34A", "bold": true}, "grid": {"gridHorizontal": true, "gridHorizontalColor": "#E5E7EB"}, "values": {"fontSize": 11}}, "*": {"title": {"fontSize": 14, "bold": true}}}}. Values plain ("#RRGGBB", numbers, true/false); the rest of the theme is kept.',
      args: z.object({
        dataColors: z.array(z.string()).optional(),
        background: z.string().optional(),
        foreground: z.string().optional(),
        tableAccent: z.string().optional(),
        styles: z.record(z.string(), z.record(z.string(), z.record(z.string(), z.unknown()))).optional().describe("visual type (pivotTable, tableEx, card, lineChart… or *) → object → property → value"),
      }),
      run: async ({ styles, ...colors }) => {
        const current = ((await r().getTheme().catch(() => undefined)) as { themeJson?: Record<string, unknown> } | undefined)?.themeJson ?? { name: "Hub theme" };
        const given = Object.fromEntries(Object.entries(colors).filter(([, v]) => v !== undefined));
        const visualStyles = Object.fromEntries(
          Object.entries(styles ?? {}).map(([type, objects]) => [type, { "*": Object.fromEntries(Object.entries(objects).map(([o, props]) => [o, [Object.fromEntries(Object.entries(props).map(([p, v]) => [p, themeValue(v)]))]])) }]),
        );
        await r().applyTheme({ themeJson: merge(current, { ...given, ...(styles && { visualStyles }) }) });
        return `Theme updated (${[...Object.keys(given), ...Object.keys(styles ?? {}).map((t) => `styles of ${t}`)].join(", ") || "nothing"}).`;
      },
    }),
    daxTool(datasetRef),
    queryTool(() => datasetRef() && { ...datasetRef()!, info: dataset()?.info }),
    ...modelTools(dataset),
  ].map(mark);

  return {
    instructions: BUILDER_EDITOR_INSTRUCTIONS,
    tools,
    // The reference reports the user chose (Reports tab) come with the context.
    context: async (focus) => [await reportContext(entry()!, dataset(), focus), await examplesText(entry()?.exampleIds)].filter(Boolean).join("\n\n"),
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
