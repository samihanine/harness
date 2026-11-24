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

export function reportActions(report: () => Report | undefined): BrowserActions {
  const need = () => {
    const r = report();
    if (!r) throw new Error("The report is not loaded on screen yet.");
    return r;
  };
  return {
    "report.state": async () => {
      const r = need();
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
      const r = need();
      const out: unknown[] = [];
      for (const p of await r.getPages()) {
        if (p.visibility === 1) continue;
        const visuals = await p.getVisuals().catch(() => []);
        for (const v of visuals.filter((x) => x.type === "slicer" || x.type === "advancedSlicerVisual")) out.push({ page: p.displayName, ...(await slicerInfo(v)) });
      }
      return out;
    },
    "report.bookmark": async () => (await need().bookmarksManager.capture()).state,
    "report.setSlicer": async ({ name, values, range, clear }: { name: string; values?: unknown[]; range?: { from?: unknown; to?: unknown }; clear?: boolean }) => {
      const { page } = await currentPage(need());
      const v = findVisual(await page.getVisuals(), name);
      if (!v) throw new Error(`No slicer "${name}" on the current page.`);
      const state = await v.getSlicerState();
      const target = state.targets?.[0] as Target | undefined;
      if (!target) throw new Error("This slicer has no field.");
      if (clear) return v.setSlicerState({ filters: [] });
      if (range) {
        const conditions = [range.from !== undefined && { operator: "GreaterThanOrEqual", value: range.from }, range.to !== undefined && { operator: "LessThanOrEqual", value: range.to }].filter(Boolean);
        return v.setSlicerState({ filters: [{ $schema: "http://powerbi.com/product/schema#advanced", target, logicalOperator: "And", conditions, filterType: 0 } as never] });
      }
      return v.setSlicerState({ filters: [{ $schema: "http://powerbi.com/product/schema#basic", target, operator: "In", values: values ?? [], filterType: 1, requireSingleSelection: false } as never] });
    },
    "report.setFilter": async ({ table, column, values, clear, level }: { table: string; column: string; values: unknown[]; clear?: boolean; level: "page" | "report" }) => {
      const r = need();
      const { pbi } = await powerbi();
      const host = level === "report" ? r : (await currentPage(r)).page;
      const current = (await host.getFilters()) as any[];
      const others = current.filter((f) => !(f.target?.table === table && f.target?.column === column));
      const next = clear ? others : [...others, { $schema: "http://powerbi.com/product/schema#basic", target: { table, column }, operator: "In", values, filterType: pbi.models.FilterType.Basic }];
      await host.updateFilters(pbi.models.FiltersOperations.ReplaceAll, next as never);
    },
    "report.setPage": async ({ page }: { page: string }) => {
      const { pages } = await currentPage(need());
      const p = pages.find((x) => x.name === page) ?? pages.find((x) => x.displayName.toLowerCase() === page.toLowerCase());
      if (!p) throw new Error(`No page "${page}". Pages: ${pages.map((x) => x.displayName).join(", ")}`);
      await p.setActive();
    },
    "report.reset": async () => {
      const r = need() as Report & { initialState?: string };
      if (r.initialState) await r.bookmarksManager.applyState(r.initialState);
      else await r.reload();
    },
    "report.visualData": async ({ visual, rows }: { visual: string; rows: number }) => {
      const { pbi } = await powerbi();
      const { page } = await currentPage(need());
      const v = findVisual(await page.getVisuals(), visual);
      if (!v) throw new Error(`No visual "${visual}" on the current page.`);
      const data = await v.exportData(pbi.models.ExportDataType.Summarized, rows);
      return { title: v.title, csv: data.data };
    },
  };
}
