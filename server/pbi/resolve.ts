/**
 * A report or a model from its ids (a link): its name, its model and where that model lives. A report of an
 * app or shared with the user is found by trying each place; its model's workspace is the report's
 * `datasetWorkspaceId` (an app's report sits in the app, its model in the app's workspace).
 */
import type { DatasetRef, Ref, ReportRef } from "@shared/types";
import { pbi, scope } from "../ms/api";
import { reportCatalog, discoverReport, discoverDataset } from "./catalog";

export async function resolveReport(r: Ref): Promise<ReportRef> {
  const known = (await reportCatalog().catch(() => [])).find((x) => x.id === r.id);
  const tries = [
    r.appId && (() => pbi(`/apps/${r.appId}/reports/${r.id}`)),
    r.groupId && (() => pbi(`/groups/${r.groupId}/reports/${r.id}`)),
    known?.appId && (() => pbi(`/apps/${known.appId}/reports/${r.id}`)),
    known?.groupId && (() => pbi(`/groups/${known.groupId}/reports/${r.id}`)),
    () => pbi(`/reports/${r.id}`),
  ].filter(Boolean) as (() => Promise<any>)[];
  let meta: any;
  for (const t of tries) if ((meta = await t().catch(() => undefined))) break;
  const at = { id: r.id, groupId: r.groupId ?? known?.groupId, appId: r.appId ?? known?.appId };
  // Not readable by the API (shared to the user only): it can still be embedded by its id.
  if (!meta) return { ...at, name: known?.name ?? `Report ${r.id.slice(0, 8)}`, datasetId: known?.datasetId, access: "view" };
  const datasetGroupId: string | undefined = meta.datasetWorkspaceId ?? (meta.datasetId ? (await pbi(`/datasets/${meta.datasetId}`).catch(() => undefined))?.workspaceId : undefined) ?? (at.appId ? undefined : at.groupId);
  if (at.appId) discoverReport({ id: r.id, name: meta.name, appId: at.appId, workspace: "App", datasetId: meta.datasetId, webUrl: meta.webUrl });
  return { ...at, name: meta.name, datasetId: meta.datasetId, datasetGroupId: datasetGroupId ?? undefined, webUrl: meta.webUrl };
}

export async function resolveDataset(r: { id: string; groupId?: string }): Promise<DatasetRef> {
  const meta = await pbi(`${scope(r.groupId)}/datasets/${r.id}`).catch(() => pbi(`/datasets/${r.id}`));
  const groupId = r.groupId ?? meta.workspaceId ?? undefined;
  discoverDataset({ id: r.id, name: meta.name, groupId, workspace: groupId ? "Workspace" : "Shared" });
  return { id: r.id, groupId, name: meta.name };
}
