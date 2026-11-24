/** The reports the signed-in user can open (My workspace and every workspace), kept 5 minutes. */
import { pbi } from "../ms/api";

export type CatalogReport = { id: string; name: string; groupId?: string; workspace: string; datasetId?: string; webUrl?: string };
let cache: { at: number; list: Promise<CatalogReport[]> } | undefined;

export function reportCatalog(): Promise<CatalogReport[]> {
  if (cache && Date.now() - cache.at < 5 * 60_000) return cache.list;
  const list = (async () => {
    const mine = (await pbi("/reports")).value.map((r: any) => ({ id: r.id, name: r.name, workspace: "My workspace", datasetId: r.datasetId, webUrl: r.webUrl }));
    const groups = (await pbi("/groups?$top=200").catch(() => ({ value: [] }))).value as { id: string; name: string }[];
    const inGroups = await Promise.all(groups.map((g) => pbi(`/groups/${g.id}/reports`).then((r) => r.value.map((x: any) => ({ id: x.id, name: x.name, groupId: g.id, workspace: g.name, datasetId: x.datasetId, webUrl: x.webUrl })), () => [])));
    return [...mine, ...inGroups.flat()] as CatalogReport[];
  })();
  cache = { at: Date.now(), list };
  list.catch(() => (cache = undefined));
  return list;
}

export async function datasetCatalog() {
  const mine = (await pbi("/datasets")).value.map((d: any) => ({ id: d.id, name: d.name, workspace: "My workspace" }));
  const groups = (await pbi("/groups?$top=200").catch(() => ({ value: [] }))).value as { id: string; name: string }[];
  const inGroups = await Promise.all(groups.map((g) => pbi(`/groups/${g.id}/datasets`).then((r) => r.value.map((x: any) => ({ id: x.id, name: x.name, groupId: g.id, workspace: g.name })), () => [])));
  return [...mine, ...inGroups.flat()] as { id: string; name: string; groupId?: string; workspace: string }[];
}
