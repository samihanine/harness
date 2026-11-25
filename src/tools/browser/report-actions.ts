/**
 * The browser side of the report tools: actions on the embedded report the server asks for (it does
 * not hold the embed). Read the state (pages, slicers, filters, visuals), change slicers, filters and
 * pages, take the view's state (a bookmark the server applies to take a picture), export a visual's data.
 */
import type * as PbiClient from "powerbi-client";
import type { BrowserActions } from "@/lib/chat";
import { powerbi, type Report } from "@/features/powerbi/embed";

type Visual = PbiClient.VisualDescriptor;
type Target = { table: string; column?: string; measure?: string; hierarchy?: string; hierarchyLevel?: string };

/** A readable message of what the Power BI client rejects with (often a plain object, not an Error). */
export function describeError(e: unknown): string {
  if (e instanceof Error) return e.message;
  if (e && typeof e === "object") {
    const o = e as Record<string, any>;
    const d = o.detail ?? o.body ?? o;
    const parts = [d?.message, d?.detailedMessage, ...(Array.isArray(d) ? d.map((x: any) => x?.message ?? JSON.stringify(x)) : []), ...(Array.isArray(d?.errors) ? d.errors.map((x: any) => x?.message ?? JSON.stringify(x)) : [])].filter((x) => typeof x === "string" && x);
    return parts.length ? [...new Set(parts)].join(": ") : JSON.stringify(e).slice(0, 600);
  }
  return String(e);
}

const fieldOf = (t?: Target) => ({ table: t?.table ?? "?", column: t?.column ?? t?.hierarchyLevel ?? t?.measure ?? t?.hierarchy ?? "?" });
const active = (filters: any[]) => filters.filter((f) => (f.values?.length && f.operator !== "All") || f.conditions?.length || f.hierarchyData?.length);
const filterText = (f: any) => {
  const t = fieldOf(f.target);
  if (f.values) return `'${t.table}'[${t.column}] ${f.operator ?? "In"} ${JSON.stringify(f.values)}`;
  if (f.conditions) return `'${t.table}'[${t.column}] ${f.conditions.map((c: any) => `${c.operator} ${c.value ?? ""}`).join(` ${f.logicalOperator} `)}`;
  if (f.hierarchyData) return `hierarchy filter on ${t.table}`;
  return `'${t.table}'[${t.column}] (${f.filterType})`;
};

async function currentPage(report: Report) {
  const pages = await report.getPages();
  return { pages, page: pages.find((p) => p.isActive) ?? pages[0] };
}

async function slicerInfo(v: Visual) {
  const s = await v.getSlicerState().catch(() => null);
  const selected = active(s?.filters ?? []);
  return { name: v.name, title: v.title || undefined, fields: (s?.targets ?? []).map((t) => fieldOf(t as Target)), selected: selected.length ? selected.map(filterText).join("; ") : "nothing selected (all)", hidden: v.layout.displayState?.mode === 1 || undefined };
}

const findVisual = (visuals: Visual[], key: string) => visuals.find((v) => v.name === key) ?? visuals.find((v) => (v.title ?? "").toLowerCase() === key.toLowerCase());

/**
 * `ready`: resolves once the report on screen has drawn (the actions wait for it; the server tells the chat
 * it is waiting for the report to load).
 */
export function reportActions(report: () => Report | undefined, ready: () => Promise<void> = async () => undefined): BrowserActions {
  const need = async () => {
    await ready();
    const r = report();
    if (!r) throw new Error("The report is not loaded on screen.");
    return r;
  };
  return {
    "report.state": async () => {
      const r = await need();
      const { pages, page } = await currentPage(r);
      const visuals = await page.getVisuals();
      return {
        pages: pages.map((p) => ({ name: p.name, displayName: p.displayName, active: p.isActive, hidden: p.visibility === 1 || undefined })),
        reportFilters: active(await r.getFilters().catch(() => [])).map(filterText),
        pageFilters: active(await page.getFilters().catch(() => [])).map(filterText),
        slicers: await Promise.all(visuals.filter((v) => v.type === "slicer" || v.type === "advancedSlicerVisual").map(slicerInfo)),
        visuals: visuals.filter((v) => v.layout.displayState?.mode !== 1).map((v) => ({ name: v.name, type: v.type, title: v.title || undefined })),
      };
    },
    "report.slicers": async () => {
      const r = await need();
      const out: unknown[] = [];
      for (const p of await r.getPages()) {
        if (p.visibility === 1) continue;
        const visuals = await p.getVisuals().catch(() => []);
        for (const v of visuals.filter((x) => x.type === "slicer" || x.type === "advancedSlicerVisual")) out.push({ page: p.displayName, ...(await slicerInfo(v)) });
      }
      return out;
    },
    "report.bookmark": async () => (await (await need()).bookmarksManager.capture()).state,
    "report.setSlicer": async ({ name, values, range, clear }: { name: string; values?: unknown[]; range?: { from?: unknown; to?: unknown }; clear?: boolean }) => {
      const r = await need();
      const { pbi } = await powerbi();
      const { page } = await currentPage(r);
      const v = findVisual(await page.getVisuals(), name);
      if (!v) throw new Error(`No slicer "${name}" on the current page.`);
      const state = await v.getSlicerState();
      const targets = (state.targets ?? []) as Target[];
      const target = targets[0];
      if (!target) throw new Error("This slicer has no field.");
      const basic = { $schema: "http://powerbi.com/product/schema#basic", target, operator: "In", values: values ?? [], filterType: pbi.models.FilterType.Basic };
      const filter = clear
        ? undefined
        : range
          ? { $schema: "http://powerbi.com/product/schema#advanced", target, logicalOperator: "And", conditions: [range.from !== undefined && { operator: "GreaterThanOrEqual", value: range.from }, range.to !== undefined && { operator: "LessThanOrEqual", value: range.to }].filter(Boolean), filterType: pbi.models.FilterType.Advanced }
          : targets.length > 1
            ? // A hierarchy slicer: its top level's values selected whole.
              { $schema: "http://powerbi.com/product/schema#hierarchy", target: targets, hierarchyData: (values ?? []).map((value) => ({ operator: "Selected", value })), filterType: pbi.models.FilterType.Hierarchy }
            : basic;
      try {
        await v.setSlicerState({ filters: filter ? [filter as never] : [] });
        return { ok: true };
      } catch (e) {
        const why = describeError(e);
        // Some slicers (new slicer types, hierarchies, sync groups) refuse a state set by the API:
        // the same selection is applied as a page filter on the slicer's column, the slicer itself unchanged.
        if (clear) throw new Error(`The slicer refused to be cleared: ${why}`);
        const fallback = targets.length > 1 ? { ...basic, target: { table: target.table, column: target.column ?? target.hierarchyLevel } } : (filter as typeof basic);
        try {
          const current = (await page.getFilters()) as any[];
          const others = current.filter((f) => !(f.target?.table === fallback.target.table && (f.target?.column ?? f.target?.hierarchyLevel) === (fallback.target as Target).column));
          await page.updateFilters(pbi.models.FiltersOperations.ReplaceAll, [...others, fallback] as never);
        } catch (e2) {
          throw new Error(`Power BI refused to change this slicer (${why}) and the same filter on the page (${describeError(e2)}). Slicer on ${targets.map((t) => `'${t.table}'[${t.column ?? t.hierarchyLevel ?? t.measure}]`).join(" > ")}; values given: ${JSON.stringify(values ?? range)}.`);
        }
        return { ok: true, note: `The slicer itself refused the selection (${why}): the same selection is applied as a page filter on its column (the slicer still shows its old selection).` };
      }
    },
    "report.setFilter": async ({ table, column, values, clear, level }: { table: string; column: string; values: unknown[]; clear?: boolean; level: "page" | "report" }) => {
      const r = await need();
      const { pbi } = await powerbi();
      const host = level === "report" ? r : (await currentPage(r)).page;
      const current = (await host.getFilters()) as any[];
      const others = current.filter((f) => !(f.target?.table === table && f.target?.column === column));
      const next = clear ? others : [...others, { $schema: "http://powerbi.com/product/schema#basic", target: { table, column }, operator: "In", values, filterType: pbi.models.FilterType.Basic }];
      await host.updateFilters(pbi.models.FiltersOperations.ReplaceAll, next as never);
    },
    "report.setPage": async ({ page }: { page: string }) => {
      const { pages } = await currentPage(await need());
      const p = pages.find((x) => x.name === page) ?? pages.find((x) => x.displayName.toLowerCase() === page.toLowerCase());
      if (!p) throw new Error(`No page "${page}". Pages: ${pages.map((x) => x.displayName).join(", ")}`);
      await p.setActive();
    },
    "report.reset": async () => {
      const r = (await need()) as Report & { initialState?: string };
      if (r.initialState) await r.bookmarksManager.applyState(r.initialState);
      else await r.reload();
    },
    "report.visualData": async ({ visual, rows }: { visual: string; rows: number }) => {
      const { pbi } = await powerbi();
      const { page } = await currentPage(await need());
      const v = findVisual(await page.getVisuals(), visual);
      if (!v) throw new Error(`No visual "${visual}" on the current page.`);
      const data = await v.exportData(pbi.models.ExportDataType.Summarized, rows);
      return { title: v.title, csv: data.data };
    },
  };
}
