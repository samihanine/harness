/** Agent tools on an embedded report (viewer: navigation & filters; builder: authoring) and on DAX. */
import { z } from "zod";
import { basicFilter, dax, models, rowsText, settled, type Ref, type Report } from "@/lib/pbi";
import { tool, type Tool } from "./loop";

export const daxTool = (dataset: () => Ref | undefined): Tool =>
  tool({
    name: "run_dax",
    description: "Runs a DAX query (EVALUATE …) on the semantic model and returns the rows (max 50 shown). Use TOPN / SUMMARIZECOLUMNS to keep results small.",
    args: z.object({ query: z.string().describe("DAX query starting with EVALUATE (or DEFINE)") }),
    readOnly: true,
    run: async ({ query }) => {
      const ref = dataset();
      if (!ref) throw new Error("No semantic model selected.");
      return rowsText(await dax(ref, query));
    },
  });

/** "Table[Column]" / "Table[Measure]" / "Sum(Table[Column])" → Power BI target. `measures` = "Table[Measure]" names. */
function target(spec: string, measures: Set<string>) {
  const m = spec.trim().match(/^(?:(\w+)\()?'?([^'[]+)'?\[([^\]]+)\]\)?$/);
  if (!m) throw new Error(`Field "${spec}" must look like Table[Column], Table[Measure] or Sum(Table[Column])`);
  const [, aggregation, table, name] = m;
  if (measures.has(`${table}[${name}]`)) return { $schema: "http://powerbi.com/product/schema#measure", table, measure: name };
  if (aggregation) return { $schema: "http://powerbi.com/product/schema#columnAggr", table, column: name, aggregationFunction: aggregation[0].toUpperCase() + aggregation.slice(1).toLowerCase() };
  return { $schema: "http://powerbi.com/product/schema#column", table, column: name };
}

const pageOf = async (report: Report, name?: string) => {
  const pages = await report.getPages();
  const page = name ? pages.find((p) => p.name === name || p.displayName === name) : pages.find((p) => p.isActive);
  if (!page) throw new Error(`Page "${name}" not found. Pages: ${pages.map((p) => p.displayName).join(", ")}`);
  return page;
};
const visualOf = async (report: Report, name: string) => {
  const visuals = await (await pageOf(report)).getVisuals();
  const visual = visuals.find((v) => v.name === name || v.title === name);
  if (!visual) throw new Error(`Visual "${name}" not on the current page.`);
  return visual;
};

/** Viewer: navigate and filter (works with view-only rights). */
export function viewTools(report: () => Report | undefined): Tool[] {
  const r = () => {
    const current = report();
    if (!current) throw new Error("No report open.");
    return current;
  };
  // The page is read again after each tool: wait for the visuals to show the change.
  const after = (t: Tool): Tool => ({ ...t, run: async (args) => {
    const rendered = settled(r());
    const text = await t.run(args);
    await rendered;
    return text;
  } });
  return [
    tool({
      name: "go_to_page",
      description: "Shows another page of the report.",
      args: z.object({ page: z.string().describe("Page name or display name") }),
      run: async ({ page }) => {
        await (await pageOf(r(), page)).setActive();
        return "Page shown.";
      },
    }),
    tool({
      name: "set_filter",
      description: "Filters the report or the current page on a column (values = kept values). Empty values removes the filter on that column.",
      args: z.object({
        level: z.enum(["report", "page"]),
        table: z.string(),
        column: z.string(),
        values: z.array(z.union([z.string(), z.number(), z.boolean()])),
      }),
      run: async ({ level, table, column, values }) => {
        const owner = level === "report" ? r() : await pageOf(r());
        const others = (await owner.getFilters()).filter((f: any) => !(f.target?.table === table && f.target?.column === column));
        await owner.setFilters(values.length ? [...others, basicFilter(table, column, values)] : others as never);
        return values.length ? `Filter set on ${table}[${column}].` : `Filter on ${table}[${column}] removed.`;
      },
    }),
    tool({
      name: "clear_filters",
      description: "Removes every filter set on the report or the current page.",
      args: z.object({ level: z.enum(["report", "page"]) }),
      run: async ({ level }) => {
        await (level === "report" ? r() : await pageOf(r())).removeFilters();
        return "Filters removed.";
      },
    }),
    tool({
      name: "set_slicer",
      description: "Selects values in a slicer of the current page (empty values clears it).",
      args: z.object({ visual: z.string().describe("Slicer name or title"), values: z.array(z.union([z.string(), z.number(), z.boolean()])) }),
      run: async ({ visual, values }) => {
        const slicer = await visualOf(r(), visual);
        const { targets } = await slicer.getSlicerState();
        const t = targets?.[0] as { table: string; column: string } | undefined;
        if (!t) throw new Error("This visual is not a slicer on a column.");
        await slicer.setSlicerState({ filters: values.length ? [basicFilter(t.table, t.column, values) as never] : [] });
        return "Slicer updated.";
      },
    }),
  ].map(after);
}

/** Builder: pages and visuals through the authoring API (edit mode). */
export function authorTools(report: () => Report | undefined, measures: () => Set<string>): Tool[] {
  const r = () => {
    const current = report();
    if (!current) throw new Error("No report open.");
    return current;
  };
  const prop = (value: unknown) => ({ schema: "http://powerbi.com/product/schema#property", value });
  return [
    ...viewTools(report),
    tool({
      name: "add_page",
      description: "Adds a page and shows it.",
      args: z.object({ displayName: z.string() }),
      run: async ({ displayName }) => {
        const page = await r().addPage(displayName);
        await page.setActive();
        return `Page added (name=${page.name}).`;
      },
    }),
    tool({
      name: "rename_page",
      description: "Renames a page.",
      args: z.object({ page: z.string(), displayName: z.string() }),
      run: async ({ page, displayName }) => {
        await r().renamePage((await pageOf(r(), page)).name, displayName);
        return "Renamed.";
      },
    }),
    tool({
      name: "delete_page",
      description: "Deletes a page.",
      args: z.object({ page: z.string() }),
      run: async ({ page }) => {
        await r().deletePage((await pageOf(r(), page)).name);
        return "Deleted.";
      },
    }),
    tool({
      name: "create_visual",
      description:
        "Creates a visual on the current page. Types: clusteredColumnChart, clusteredBarChart, lineChart, areaChart, pieChart, donutChart, card, multiRowCard, tableEx, pivotTable, slicer, scatterChart, gauge, treemap, map, kpi, textbox. Page is 1280×720.",
      args: z.object({
        type: z.string(),
        x: z.number(),
        y: z.number(),
        width: z.number(),
        height: z.number(),
        fields: z
          .record(z.string(), z.array(z.string()))
          .optional()
          .describe('By role, e.g. {"Category": ["Account[Region]"], "Y": ["Sum(Fact[Amount])"]} (roles: Category, Y, Series, Values, Rows, Columns, Tooltips…)'),
        title: z.string().optional(),
      }),
      run: async ({ type, x, y, width, height, fields, title }) => {
        const page = await pageOf(r());
        const { visual } = await page.createVisual(type, { x, y, width, height });
        for (const [role, list] of Object.entries(fields ?? {})) for (const f of list) await visual.addDataField(role, target(f, measures()) as never);
        if (title) await visual.setProperty({ objectName: "title", propertyName: "titleText" }, prop(title) as never);
        // Created visuals can come out hidden: make sure it is shown.
        await page.setVisualDisplayState(visual.name, (await models()).VisualContainerDisplayMode.Visible).catch(() => undefined);
        return `Visual created (name=${visual.name}).`;
      },
    }),
    tool({
      name: "set_visual_fields",
      description: "Replaces the fields of one role of a visual.",
      args: z.object({ visual: z.string(), role: z.string(), fields: z.array(z.string()) }),
      run: async ({ visual, role, fields }) => {
        const v = await visualOf(r(), visual);
        for (let i = (await v.getDataFields(role)).length - 1; i >= 0; i--) await v.removeDataField(role, i);
        for (const f of fields) await v.addDataField(role, target(f, measures()) as never);
        return "Fields updated.";
      },
    }),
    tool({
      name: "format_visual",
      description: 'Sets a format property, e.g. ("title","titleText","Sales"), ("title","fontColor","#C00000"), ("background","color","#FFF2CC"), ("legend","position","Top"), ("legend","visible",false), ("labels","visible",true). Series colors are set with set_theme.',
      args: z.object({ visual: z.string(), object: z.string(), property: z.string(), value: z.union([z.string(), z.number(), z.boolean()]) }),
      run: async ({ visual, object, property, value }) => {
        await (await visualOf(r(), visual)).setProperty({ objectName: object, propertyName: property }, prop(value) as never);
        return "Formatted.";
      },
    }),
    tool({
      name: "change_visual",
      description: "Changes the type and/or position/size of a visual.",
      args: z.object({ visual: z.string(), type: z.string().optional(), x: z.number().optional(), y: z.number().optional(), width: z.number().optional(), height: z.number().optional() }),
      run: async ({ visual, type, x, y, width, height }) => {
        const page = await pageOf(r());
        const v = await visualOf(r(), visual);
        if (type) await v.changeType(type);
        if (x !== undefined || y !== undefined) await page.moveVisual(v.name, x ?? v.layout.x ?? 0, y ?? v.layout.y ?? 0);
        if (width !== undefined || height !== undefined) await page.resizeVisual(v.name, width ?? v.layout.width ?? 300, height ?? v.layout.height ?? 200);
        return "Changed.";
      },
    }),
    tool({
      name: "delete_visual",
      description: "Deletes a visual of the current page.",
      args: z.object({ visual: z.string() }),
      run: async ({ visual }) => {
        await (await pageOf(r())).deleteVisual((await visualOf(r(), visual)).name);
        return "Deleted.";
      },
    }),
    tool({
      name: "set_theme",
      description: "Applies a theme to the whole report (series colors, background, text color).",
      args: z.object({ dataColors: z.array(z.string()).describe("Hex colors of the series"), background: z.string().optional(), foreground: z.string().optional() }),
      run: async (theme) => {
        await r().applyTheme({ themeJson: { name: "custom", ...theme } });
        return "Theme applied.";
      },
    }),
    tool({
      name: "save_report",
      description: "Saves the report (the user's changes and yours). Call it when a requested change is done.",
      args: z.object({}),
      run: async () => {
        await r().save();
        return "Saved.";
      },
    }),
  ];
}
