/** Agent tools of the report viewer (navigation, filters, slicers, bookmarks, data of a visual) and DAX. */
import { z } from "zod";
import { basicFilter, checkValues, columnValues, dax, findVisual, models, rowsText, settled, visualText, type Ref, type Report } from "@/lib/pbi";
import { checkName, type ModelInfo } from "@/lib/model";
import { tool, type Tool } from "./loop";

/** How to write DAX that runs the first time (given to every agent writing queries). */
export const DAX_RULES = `DAX rules (follow them, queries fail otherwise):
- Use ONLY tables, columns and measures that exist, with their exact names copied from the model (context, describe_table). Never guess, shorten or translate a name. Not sure a column exists or in which table it is? search_model first.
- Syntax: 'Table'[Column], [Measure], text in "double quotes". Count rows with COUNTROWS('Table') or an existing measure.
- Grouping: SUMMARIZECOLUMNS('T'[a], 'T'[b], "Label", [Measure]). A plain list of value pairs of ONE table: DISTINCT(SELECTCOLUMNS('T', "A", 'T'[a], "B", 'T'[b])). Columns of two tables together only if they are related (see the relationships).
- ORDER BY only columns of the result. Count the parentheses before sending.
- A failed query: read the error, it names the problem (a missing name, or a (line, column) position) — fix exactly that and retry. To explain a failure to the user, quote the error; never invent a cause.`;

/** run_dax: on the screen's model, or (`several`) on one of its models chosen by id. A thinking tool: results are capped. */
export const daxTool = (dataset: () => Ref | undefined, several?: () => (Ref & { name: string })[]): Tool =>
  tool({
    name: "run_dax",
    description: `Runs a DAX query (EVALUATE …) on ${several ? "one of the semantic models (dataset_id)" : "the semantic model"} and returns the rows (max 50 shown). Use it to check values, totals or names before acting. Use TOPN / SUMMARIZECOLUMNS to keep results small. Exact names only (see the DAX rules).`,
    args: z.object({ query: z.string().describe("DAX query starting with EVALUATE (or DEFINE)"), dataset_id: z.string().optional().describe("Model id (when there are several)") }),
    readOnly: true,
    run: async ({ query, dataset_id }) => {
      const list = several?.() ?? [];
      const ref = dataset_id ? list.find((d) => d.id === dataset_id || d.name === dataset_id) : (dataset() ?? (list.length === 1 ? list[0] : undefined));
      if (!ref) throw new Error(list.length ? `Give dataset_id: one of ${list.map((d) => `${d.id} (${d.name})`).join(", ")}.` : "No semantic model selected.");
      return rowsText(await dax(ref, query));
    },
  });

const pageOf = async (report: Report, name?: string) => {
  const pages = await report.getPages();
  const page = name ? pages.find((p) => p.name === name || p.displayName === name) : pages.find((p) => p.isActive);
  if (!page) throw new Error(`Page "${name}" not found. Pages: ${pages.map((p) => p.displayName).join(", ")}`);
  return page;
};
/** By name (id) or title ("Selected item" = "selected_item"); a title shared by several visuals needs the name. */
const visualOf = async (report: Report, name: string) => findVisual(await (await pageOf(report)).getVisuals(), name);
/** One value or a list. */
const values = z.preprocess((v) => (Array.isArray(v) ? v : v === undefined || v === null || v === "" ? [] : [v]), z.array(z.union([z.string(), z.number(), z.boolean()])));

/**
 * Selection of a hierarchy slicer (several levels: a hierarchy, or several columns) from paths
 * "top > sub > …": each part is checked against its level's values (case fixed), then the paths become
 * the slicer's tree — the last part of a path selected, the parts above it "inherited".
 */
async function hierarchyFilter(levels: { table: string; column?: string; hierarchy?: string; hierarchyLevel?: string }[], paths: string[], dataset?: Ref) {
  type Node = { operator: "Selected" | "Inherited"; value: unknown; children?: Node[] };
  const roots: Node[] = [];
  for (const path of paths) {
    const parts = path.split(/\s*>\s*/).filter(Boolean);
    if (parts.length > levels.length) throw new Error(`"${path}" has ${parts.length} parts; this slicer has ${levels.length} levels.`);
    let siblings = roots;
    for (const [i, part] of parts.entries()) {
      const level = levels[i];
      const column = level.column ?? level.hierarchyLevel ?? "";
      // The value as the model holds it (a year is a number, not "2024"), else checked (refused with the valid ones).
      const known = dataset ? await columnValues(dataset, level.table, column).catch(() => undefined) : undefined;
      let value = known?.values.find((v) => String(v).toLowerCase() === part.toLowerCase());
      if (value === undefined) [value] = await checkValues(dataset, level.table, column, [/^-?\d+(\.\d+)?$/.test(part) ? Number(part) : part]);
      let node = siblings.find((n) => String(n.value).toLowerCase() === String(value).toLowerCase());
      if (!node) siblings.push((node = { operator: "Inherited", value }));
      if (i === parts.length - 1) {
        node.operator = "Selected";
        delete node.children;
        break;
      }
      if (node.operator === "Selected") break; // already selected whole
      siblings = node.children ??= [];
    }
  }
  return {
    $schema: "http://powerbi.com/product/schema#hierarchy",
    target: levels.map((l) => (l.column ? { table: l.table, column: l.column } : { table: l.table, hierarchy: l.hierarchy, hierarchyLevel: l.hierarchyLevel })),
    filterType: (await models()).FilterType.Hierarchy,
    hierarchyData: roots,
  };
}

/**
 * Report viewer: change the page, filters, slicers and bookmarks, read a visual's data (works with view
 * rights). With the model (`dataset`), filter values are checked first: an unknown value is refused with
 * the valid ones, instead of emptying the visuals.
 */
export function viewerTools(report: () => Report | undefined, dataset: () => Ref | undefined = () => undefined, info: () => ModelInfo | undefined = () => undefined): Tool[] {
  const r = () => {
    const current = report();
    if (!current) throw new Error("No report open.");
    return current;
  };
  // The page is read again after each change: wait for the visuals to show it.
  const after = (t: Tool): Tool => (t.readOnly ? t : { ...t, run: async (args) => {
    const rendered = settled(r());
    const text = await t.run(args);
    await rendered;
    return text;
  } });
  /** A column filter (values kept) on a filter owner (report, page, visual): replaces the one on that column; no values removes it. */
  const columnFilter = async (owner: { getFilters: () => Promise<unknown[]>; setFilters: (f: never) => Promise<unknown> }, t0: string, c0: string, given: unknown[]) => {
    const { table, name: column } = checkName(info(), t0, c0);
    const values = await checkValues(dataset(), table, column, given);
    const others = (await owner.getFilters()).filter((f: any) => !(f.target?.table === table && f.target?.column === column));
    await owner.setFilters((values.length ? [...others, basicFilter(table, column, values)] : others) as never);
    return values.length ? `Filter on ${table}[${column}] = ${values.map((v) => JSON.stringify(v)).join(", ")}.` : `Filter on ${table}[${column}] removed.`;
  };
  const filterArgs = { table: z.string(), column: z.string(), values: values.describe("Values kept, as they are in the data (empty: removes the filter on that column)") };
  return [
    tool({
      name: "change_current_page",
      description: "Shows another page of the report.",
      args: z.object({ page: z.string().describe("Page name (id) or display name") }),
      run: async ({ page }) => {
        await (await pageOf(r(), page)).setActive();
        return "Page shown.";
      },
    }),
    tool({
      name: "get_visual_data",
      description: "The data a visual of the current page shows now (first rows, with the filters applied), its fields and filters.",
      args: z.object({ visual: z.string().describe("Visual name (id), or its title when no other visual has it"), rows: z.number().optional() }),
      readOnly: true,
      run: async ({ visual }) => visualText(await visualOf(r(), visual), true, false, dataset()),
    }),
    tool({
      name: "update_report_filter",
      description: "Filters every page of the report on a column. Prefer a slicer of the page when there is one on that column.",
      args: z.object(filterArgs),
      run: async ({ table, column, values }) => columnFilter(r(), table, column, values),
    }),
    tool({
      name: "update_page_filter",
      description: "Filters the current page on a column. Prefer a slicer of the page when there is one on that column.",
      args: z.object(filterArgs),
      run: async ({ table, column, values }) => columnFilter(await pageOf(r()), table, column, values),
    }),
    tool({
      name: "update_visual_filter",
      description:
        'Filters ONE visual (others are not affected). kind "values": keeps the given values of a column. kind "top": keeps the top / bottom N items of a column ordered by another column or measure (e.g. the most recent task: column = task title, order_by = \'Tasks\'[due date], n = 1).',
      args: z.object({
        visual: z.string().describe("Visual name (id), or its title when no other visual has it"),
        kind: z.enum(["values", "top"]),
        ...filterArgs,
        values: filterArgs.values.optional(),
        n: z.number().optional(),
        order_by: z.string().optional().describe("Table[Column] (its max is used) or Table[Measure]"),
        bottom: z.boolean().optional().describe("Lowest N instead of highest"),
      }),
      run: async ({ visual, kind, table: t0, column: c0, values: kept = [], n = 1, order_by, bottom }) => {
        const v = await visualOf(r(), visual);
        if (kind === "values") return `${v.name}: ${await columnFilter(v, t0, c0, kept)}`;
        const { table, name: column } = checkName(info(), t0, c0);
        const parsed = order_by?.match(/^'?([^'[]+)'?\[([^\]]+)\]$/);
        if (!parsed) throw new Error("order_by must look like Table[Column] or Table[Measure]");
        const by = checkName(info(), parsed[1], parsed[2]);
        const others = (await v.getFilters()).filter((f: any) => !(f.target?.table === table && f.target?.column === column));
        const filter = { $schema: "http://powerbi.com/product/schema#topN", target: { table, column }, operator: bottom ? "Bottom" : "Top", itemCount: n, orderBy: { table: by.table, column: by.name, aggregationFunction: "Max" }, filterType: 5 };
        await v.setFilters([...others, filter] as never);
        return `${v.name} now shows the ${bottom ? "bottom" : "top"} ${n} ${column} by ${order_by}.`;
      },
    }),
    tool({
      name: "update_slicer_value",
      description: 'Selects values in a slicer of the current page (empty values clears it). Values as listed in the slicer\'s options; unknown values are refused with the valid ones. Hierarchy slicer (several levels, e.g. Year > Month): paths "top > sub" (e.g. "2024 > Jan"); a top value alone selects it whole.',
      args: z.object({ visual: z.string().describe("Slicer name (id), or its title when no other visual has it"), values }),
      run: async ({ visual, values }) => {
        const slicer = await visualOf(r(), visual);
        const { targets } = await slicer.getSlicerState();
        const levels = (targets ?? []) as { table: string; column?: string; hierarchy?: string; hierarchyLevel?: string }[];
        if (!levels.length) throw new Error("This visual is not a slicer on a field.");
        if (!values.length) {
          await slicer.setSlicerState({ filters: [] });
          return "Slicer cleared.";
        }
        if (levels.length === 1 && levels[0].column) {
          const t = levels[0] as { table: string; column: string };
          const kept = await checkValues(dataset(), t.table, t.column, values);
          await slicer.setSlicerState({ filters: [basicFilter(t.table, t.column, kept) as never] });
          return `Slicer set to ${kept.map((v) => JSON.stringify(v)).join(", ")}.`;
        }
        await slicer.setSlicerState({ filters: [(await hierarchyFilter(levels, values.map(String), dataset())) as never] });
        return `Slicer set to ${values.map((v) => JSON.stringify(String(v))).join(", ")}.`;
      },
    }),
    tool({
      name: "reset_filters",
      description: "Removes the filters set on the report, the current page, or both (slicers keep their selection: clear them with update_slicer_value).",
      args: z.object({ level: z.enum(["report", "page", "both"]) }),
      run: async ({ level }) => {
        if (level !== "page") await r().removeFilters();
        if (level !== "report") await (await pageOf(r())).removeFilters();
        return "Filters removed.";
      },
    }),
    tool({
      name: "apply_bookmark",
      description: "Applies a bookmark of the report (see the bookmarks listed in the page state).",
      args: z.object({ bookmark: z.string().describe("Bookmark name (id) or display name") }),
      run: async ({ bookmark }) => {
        const all = (await r().bookmarksManager.getBookmarks()).flatMap((b) => [b, ...(b.children ?? [])]);
        const b = all.find((x) => x.name === bookmark || x.displayName?.toLowerCase() === bookmark.toLowerCase());
        if (!b) throw new Error(`No bookmark "${bookmark}". Bookmarks: ${all.map((x) => `${x.displayName} (${x.name})`).join(", ") || "none"}.`);
        await r().bookmarksManager.apply(b.name);
        return `Bookmark "${b.displayName}" applied.`;
      },
    }),
  ].map(after);
}

/** The report's bookmarks, one line, for the page state. */
export async function bookmarksText(report: Report) {
  const all = (await report.bookmarksManager.getBookmarks().catch(() => [])).flatMap((b) => [b, ...(b.children ?? [])]).filter((b) => !b.children?.length);
  return all.length ? `bookmarks: ${all.map((b) => `${b.displayName} (name=${b.name})`).join(", ")}` : "bookmarks: none";
}
