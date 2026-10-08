/** Library entries: what is read from Power BI / SharePoint when a link is added or refreshed. */
import { resolveExcel } from "./excel";
import { pbi } from "./ms";
import { embedReport, modelText, parseDatasetUrl, parseReportUrl, reportText, resetEmbed, scope, type Ref } from "./pbi";
import { find, patch, put, type DatasetEntry, type ExcelEntry, type ReportEntry } from "./store";

/** Reads a report through a hidden embed: works with Viewer / Build rights. */
async function snapshot(ref: Ref) {
  const element = document.createElement("div");
  element.style.cssText = "position:fixed;left:-10000px;top:0;width:1280px;height:720px";
  document.body.append(element);
  try {
    return await reportText(await embedReport(element, ref));
  } finally {
    await resetEmbed(element);
    element.remove();
  }
}

export async function addReport(link: string): Promise<ReportEntry> {
  const ref = parseReportUrl(link);
  const r = await pbi(`${scope(ref.groupId)}/reports/${ref.id}`);
  const old = await find("reports", ref.id);
  const entry: ReportEntry = {
    ...old,
    id: ref.id,
    groupId: ref.groupId,
    name: r.name,
    url: r.webUrl ?? link,
    datasetId: r.datasetId,
    datasetGroupId: r.datasetWorkspaceId ?? ref.groupId,
    editable: r.isOwnedByMe ?? undefined,
    context: old?.context ?? "",
    snapshot: await snapshot(ref).catch((e) => `Not readable: ${e instanceof Error ? e.message : e}`),
  };
  await put("reports", entry);
  if (!(await find("datasets", r.datasetId))) await addDataset(r.datasetId, entry.datasetGroupId).catch(() => undefined);
  return entry;
}

export async function addDataset(linkOrId: string, groupId?: string, values: Partial<DatasetEntry> = {}): Promise<DatasetEntry> {
  const ref = groupId ? { id: linkOrId, groupId } : parseDatasetUrl(linkOrId);
  const d = await pbi(`${scope(ref.groupId)}/datasets/${ref.id}`);
  const old = await find("datasets", ref.id);
  return put("datasets", { ...old, id: ref.id, groupId: ref.groupId, name: d.name, context: old?.context ?? "", ...values, model: await modelText(ref) });
}

export async function refreshDatasetText(entry: DatasetEntry) {
  await patch("datasets", entry.id, { model: await modelText(entry) });
}

export async function addExcel(link: string, table?: string): Promise<ExcelEntry> {
  const file = await resolveExcel(link);
  if (file.tables.length === 0) throw new Error("This file has no Excel table: select the data in Excel and use Insert › Table.");
  return put("excels", {
    id: crypto.randomUUID(),
    name: file.name.replace(/\.xlsx$/i, ""),
    link,
    driveId: file.driveId,
    itemId: file.itemId,
    fileUrl: file.fileUrl,
    table: table && file.tables.includes(table) ? table : file.tables[0],
    context: "",
  });
}

/** Copy of a report in My workspace (needs edit rights on the source). */
export async function cloneReport(entry: ReportEntry) {
  const copy = await pbi(`${scope(entry.groupId)}/reports/${entry.id}/Clone`, "POST", {
    name: `${entry.name} (copy)`,
    targetWorkspaceId: "00000000-0000-0000-0000-000000000000",
    targetModelId: entry.datasetId,
  });
  return addReport(copy.webUrl ?? copy.id);
}

/** Power BI Pro: 8 scheduled + API refreshes per model and per UTC day; "Refresh now" clicks in Power BI are not counted. */
export const API_REFRESHES_PER_DAY = 8;

export async function apiRefreshesToday(d: Ref) {
  const today = new Date().toISOString().slice(0, 10);
  const { value } = await pbi(`${scope(d.groupId)}/datasets/${d.id}/refreshes?$top=60`);
  return (value as { refreshType: string; startTime: string }[]).filter((r) => r.refreshType !== "OnDemand" && r.startTime.startsWith(today)).length;
}

/** Semantic model page in Power BI (its "Refresh now" button is not limited). */
export const modelPage = (d: Ref) => `https://app.powerbi.com/groups/${d.groupId ?? "me"}/datasets/${d.id}/details`;
