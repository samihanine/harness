/** Agent tools of the report viewer (navigation, filters, slicers, bookmarks, data of a visual) and DAX. */
import { z } from "zod";
import { basicFilter, checkValues, columnValues, dax, findVisual, models, rowsText, settled, visualText, type Ref, type Report } from "@/services/pbi";
import { checkName, type ModelInfo } from "@/services/model";
import { QUERY_HELP, queryDax, type DefinedMeasure } from "@/lib/query-spec";
import { tool, type Tool } from "./loop";
import { PERFORMANCE } from "@/config";
import { find } from "@/lib/store";

export { DAX_RULES } from "@/prompts/dax";

export { daxLimits } from "@/prompts/dax";
import { daxLimits } from "@/prompts/dax";

/** Report measures (`defined`) put in a query's DEFINE: they may use each other. */
export function withDefined(query: string, defined: DefinedMeasure[]) {
  const defs = defined.filter((m) => m.expression).map((m) => `MEASURE '${m.table.replace(/'/g, "''")}'[${m.name}] = ${m.expression}`).join("\n");
  return !defs ? query : /^\s*DEFINE\b/i.test(query) ? query.replace(/^\s*DEFINE\b/i, `DEFINE\n${defs}`) : `DEFINE\n${defs}\n${query}`;
}

/**
 * A query failed on a name that does not exist (often a label the user pasted, e.g. a report's column
 * header): the model's closest measures and columns are added to the error, so the AI copies a real one.
 */
async function withNameHints(datasetId: string, error: unknown): Promise<Error> {
  const message = error instanceof Error ? error.message : String(error);
  const bad = message.match(/'([^']+)' cannot be determined|(?:column|measure|table) '([^']+)'|cannot find[^']*'([^']+)'|\[([^\]]+)\][^.]*(?:not found|cannot be found|doesn't exist)/i);
  const name = (bad?.[1] ?? bad?.[2] ?? bad?.[3] ?? bad?.[4])?.toLowerCase();
  const info = name ? (await find("datasets", datasetId).catch(() => undefined))?.info : undefined;
  if (!name || !info) return error instanceof Error ? error : new Error(message);
  const words = name.split(/[^\p{L}\p{N}%]+/u).filter((w) => w.length > 1);
  const names = info.tables.flatMap((t) => [...t.measures.map((m) => `[${m.name}]`), ...t.columns.map((c) => `'${t.name}'[${c.name}]`)]);
  const best = names
    .map((n) => ({ n, score: words.filter((w) => n.toLowerCase().includes(w)).length }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score || a.n.length - b.n.length)
    .slice(0, 8)
    .map((x) => x.n);
  return new Error(best.length ? `${message}\nClosest names in the model (copy one exactly): ${best.join(", ")}.` : message);
}

/**
 * run_dax: on the screen's model, or (`several`) on one of its models chosen by id; `defined`: the
 * report measures of that model, usable in the query. A thinking tool: results are capped.
 */
export const daxTool = (dataset: () => Ref | undefined, several?: () => (Ref & { name: string })[], defined?: (dataset: Ref) => Promise<DefinedMeasure[]>): Tool =>
  tool({
    name: "run_dax",
    description: `Runs a DAX query (EVALUATE …) on ${several ? "one of the semantic models (dataset_id)" : "the semantic model"} and returns the rows (max 50 shown). Use it to check values, totals or names before acting. Use TOPN / SUMMARIZECOLUMNS to keep results small. Exact names only (see the DAX rules). ${daxLimits()}`,
    args: z.object({ query: z.string().describe("DAX query starting with EVALUATE (or DEFINE)"), dataset_id: z.string().optional().describe("Model id (when there are several)") }),
    readOnly: true,
    run: async ({ query, dataset_id }) => {
      const list = several?.() ?? [];
      const ref = dataset_id ? list.find((d) => d.id === dataset_id || d.name === dataset_id) : (dataset() ?? (list.length === 1 ? list[0] : undefined));
      if (!ref) throw new Error(list.length ? `Give dataset_id: one of ${list.map((d) => `${d.id} (${d.name})`).join(", ")}.` : "No semantic model selected.");
      try {
        return rowsText(await dax(ref, defined ? withDefined(query, await defined(ref)) : query));
      } catch (error) {
        throw await withNameHints(ref.id, error);
      }
    },
  });

const scalar = z.union([z.string(), z.number(), z.boolean()]);
/** A list given as one value, or as "a, b" (weak models do both). */
const list = z.preprocess((v) => (typeof v === "string" ? v.split(/,(?![^[]*\])/).map((x) => x.trim()).filter(Boolean) : v), z.array(z.string()));
export const querySpecArgs = {
  values: list.default([]).describe("Measures [M] (model or report measures) or Sum('T'[col]) / Avg / Count / DistinctCount / Min / Max"),
  group_by: list.optional().describe("Columns 'T'[col] to group by (rows of the result)"),
  filters: z.array(z.object({ field: z.string().describe("A column 'T'[col]"), op: z.enum(["in", "not_in", ">", ">=", "<", "<=", "between", "contains"]).optional().describe("Default in"), values: z.preprocess((v) => (v === undefined || Array.isArray(v) ? v : [v]), z.array(scalar)).optional().describe("in / not_in: the values (dates YYYY-MM-DD)"), value: z.union([z.string(), z.number()]).optional(), to: z.union([z.string(), z.number()]).optional().describe("between: the upper bound") })).optional(),
  top: z.object({ count: z.number(), by: z.string().optional().describe("A value of the result (default: the first)"), ascending: z.boolean().optional().describe("true: the bottom N") }).optional(),
  order_by: z.array(z.object({ field: z.string(), descending: z.boolean().optional() })).optional(),
};

/**
 * query: a spec written as DAX by the app (query-spec.ts), on the screen's model or (`several`) one of
 * its models; `defined`: the report's own measures, usable as values.
 */
export const queryTool = (dataset: () => (Ref & { info?: ModelInfo }) | undefined, several?: () => (Ref & { name: string; info?: ModelInfo })[], defined?: (dataset: Ref) => Promise<DefinedMeasure[]>): Tool =>
  tool({
    name: "query",
    description: `${QUERY_HELP} Returns the rows (max 50 shown) and the DAX it ran. ${daxLimits()}`,
    args: z.object({ ...querySpecArgs, ...(several ? { dataset_id: z.string().optional().describe("Model id (when there are several)") } : {}) }),
    readOnly: true,
    run: async ({ dataset_id, ...spec }: any) => {
      const list = several?.() ?? [];
      const d = dataset_id ? list.find((x) => x.id === dataset_id || x.name === dataset_id) : (dataset() ?? (list.length === 1 ? list[0] : undefined));
      if (!d) throw new Error(list.length ? `Give dataset_id: one of ${list.map((x) => `${x.id} (${x.name})`).join(", ")}.` : "No semantic model selected.");
      const query = queryDax(spec, d.info, (await defined?.(d)) ?? []);
      try {
        return `${rowsText(await dax(d, query))}

DAX used:
${query}`;
      } catch (error) {
        throw new Error(`${error instanceof Error ? error.message : String(error)}
DAX used:
${query}`);
      }
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
  /**
   * A slicer of the current page on that column: the filter must go through it (the user sees and
   * changes a slicer; a hidden page filter would contradict it). Refused with the slicer to use.
   */
  const noSlicerOn = async (t0: string, c0: string) => {
    const { table, name: column } = checkName(info(), t0, c0);
    for (const v of await (await pageOf(r())).getVisuals()) {
      if (!/slicer/i.test(v.type)) continue;
      const targets = ((await v.getSlicerState().catch(() => undefined))?.targets ?? []) as { table?: string; column?: string; hierarchyLevel?: string }[];
      if (targets.some((t) => t.table?.toLowerCase() === table.toLowerCase() && (t.column ?? t.hierarchyLevel)?.toLowerCase() === column.toLowerCase()))
        throw new Error(`The slicer "${v.title ?? v.name}" (name=${v.name}) of this page is on '${table}'[${column}]: use update_slicer_value on it instead of a filter.`);
    }
  };
  return [
    tool({
      name: "list_pages",
      description: "Every page of the report with its visuals (type, title, name), the field of each slicer and its filters: to find the page (and the slicer) that answers a question before moving to it.",
      args: z.object({}),
      readOnly: true,
      run: async () => {
        const pages = await r().getPages();
        // The report's kept schema knows every slicer's field (the SDK tells it only on the current page).
        const schema = (await find("reports", r().getId()).catch(() => undefined))?.schema;
        const lines = await Promise.all(
          pages.map(async (p) => {
            const visuals = await p.getVisuals().catch(() => []);
            const filters = await p.getFilters().catch(() => []);
            const shown = visuals.filter((v) => v.layout.displayState?.mode !== 1 && !["shape", "basicShape", "image"].includes(v.type));
            // The field each slicer filters: where the user's filter can go.
            const fieldOf = async (v: (typeof shown)[number]) => {
              if (v.type !== "slicer") return "";
              const kept = schema?.visuals.find((x) => x.name === v.name && x.page === p.name);
              const fields = kept ? Object.values(kept.fields).flat() : [];
              if (fields.length) return ` on ${fields.join(", ")}`;
              const targets = ((await v.getSlicerState().catch(() => undefined))?.targets ?? []) as { table?: string; column?: string; hierarchyLevel?: string }[];
              return targets.length ? ` on ${targets.map((t) => `'${t.table}'[${t.column ?? t.hierarchyLevel}]`).join(", ")}` : "";
            };
            const described = await Promise.all(shown.map(async (v) => `${v.type} "${v.title ?? ""}" (name=${v.name})${await fieldOf(v)}`));
            return `- "${p.displayName}" (name=${p.name}${p.isActive ? ", current" : ""}${p.visibility === 1 ? ", hidden" : ""}): ${described.join("; ") || "no visual"}${filters.length ? ` — ${filters.length} page filter(s)` : ""}`;
          }),
        );
        return lines.join("\n");
      },
    }),
    tool({
      name: "read_page",
      description: "A page in detail without showing it: its visuals with their fields, filters and slicer selections (data: true adds each visual's first rows; slower).",
      args: z.object({ page: z.string().describe("Page name (id) or display name"), data: z.boolean().optional() }),
      readOnly: true,
      run: async ({ page, data = false }) => {
        const p = await pageOf(r(), page);
        const visuals = await p.getVisuals();
        const parts = await Promise.all(visuals.map((v) => visualText(v, data, false, dataset())));
        return [`page "${p.displayName}" (name=${p.name}${p.isActive ? ", current" : ""})`, ...parts].join("\n");
      },
    }),
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
      description: "Filters every page of the report on a column — only when no slicer of the page is on that column (refused then: use the slicer).",
      args: z.object(filterArgs),
      run: async ({ table, column, values }) => (await noSlicerOn(table, column), columnFilter(r(), table, column, values)),
    }),
    tool({
      name: "update_page_filter",
      description: "Filters the current page on a column — only when no slicer of the page is on that column (refused then: use the slicer).",
      args: z.object(filterArgs),
      run: async ({ table, column, values }) => (await noSlicerOn(table, column), columnFilter(await pageOf(r()), table, column, values)),
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
