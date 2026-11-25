/**
 * What the signed-in user can open: the reports and models of My workspace, of every workspace, and of
 * the Power BI apps they were given (view only), kept 5 minutes. Items opened from a pasted link are added.
 */
import { pbi } from "../ms/api";

export type CatalogReport = { id: string; name: string; groupId?: string; appId?: string; workspace: string; datasetId?: string; webUrl?: string };
export type CatalogDataset = { id: string; name: string; groupId?: string; workspace: string };

let reports: { at: number; list: Promise<CatalogReport[]> } | undefined;
let datasets: { at: number; list: Promise<CatalogDataset[]> } | undefined;
/** Found from a pasted link (an app's report or model not listed otherwise). */
const discovered = { reports: new Map<string, CatalogReport>(), datasets: new Map<string, CatalogDataset>() };

async function appReports(): Promise<CatalogReport[]> {
  const apps = ((await pbi("/apps").catch(() => ({ value: [] }))).value ?? []) as { id: string; name: string }[];
  const lists = await Promise.all(apps.map((a) => pbi(`/apps/${a.id}/reports`).then((r) => (r.value ?? []).map((x: any) => ({ id: x.id, name: x.name, appId: a.id, workspace: `App: ${a.name}`, datasetId: x.datasetId, webUrl: x.webUrl })), () => [])));
  return lists.flat();
}

export function reportCatalog(): Promise<CatalogReport[]> {
  if (!reports || Date.now() - reports.at > 5 * 60_000) {
    const list = (async () => {
      const mine = (await pbi("/reports")).value.map((r: any) => ({ id: r.id, name: r.name, workspace: "My workspace", datasetId: r.datasetId, webUrl: r.webUrl }));
      const groups = (await pbi("/groups?$top=500").catch(() => ({ value: [] }))).value as { id: string; name: string }[];
      const inGroups = await Promise.all(groups.map((g) => pbi(`/groups/${g.id}/reports`).then((r) => r.value.map((x: any) => ({ id: x.id, name: x.name, groupId: g.id, workspace: g.name, datasetId: x.datasetId, webUrl: x.webUrl })), () => [])));
      return [...mine, ...inGroups.flat(), ...(await appReports())] as CatalogReport[];
    })();
    reports = { at: Date.now(), list };
    list.catch(() => (reports = undefined));
  }
  return reports.list.then((l) => [...l, ...[...discovered.reports.values()].filter((d) => !l.some((x) => x.id === d.id))]);
}

export function datasetCatalog(): Promise<CatalogDataset[]> {
  if (!datasets || Date.now() - datasets.at > 5 * 60_000) {
    const list = (async () => {
      const mine = (await pbi("/datasets")).value.map((d: any) => ({ id: d.id, name: d.name, workspace: "My workspace" }));
      const groups = (await pbi("/groups?$top=500").catch(() => ({ value: [] }))).value as { id: string; name: string }[];
      const inGroups = await Promise.all(groups.map((g) => pbi(`/groups/${g.id}/datasets`).then((r) => r.value.map((x: any) => ({ id: x.id, name: x.name, groupId: g.id, workspace: g.name })), () => [])));
      const known = [...mine, ...inGroups.flat()] as CatalogDataset[];
      // The models behind the apps' reports (an app does not list its models).
      const fromApps = (await appReports().catch(() => [])).filter((r) => r.datasetId && !known.some((k) => k.id === r.datasetId));
      const seen = new Set<string>();
      const appModels = await Promise.all(
        fromApps
          .filter((r) => !seen.has(r.datasetId!) && seen.add(r.datasetId!))
          .map((r) => pbi(`/datasets/${r.datasetId}`).then((d) => ({ id: r.datasetId!, name: d.name ?? `${r.name} (model)`, workspace: r.workspace }), () => ({ id: r.datasetId!, name: `${r.name} (model)`, workspace: r.workspace }))),
      );
      return [...known, ...appModels];
    })();
    datasets = { at: Date.now(), list };
    list.catch(() => (datasets = undefined));
  }
  return datasets.list.then((l) => [...l, ...[...discovered.datasets.values()].filter((d) => !l.some((x) => x.id === d.id))]);
}

export const discoverReport = (r: CatalogReport) => discovered.reports.set(r.id, r);
export const discoverDataset = (d: CatalogDataset) => discovered.datasets.set(d.id, d);
