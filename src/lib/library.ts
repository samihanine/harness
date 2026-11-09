/**
 * Library entries: what is read from Power BI / SharePoint when a link is added or read again,
 * and what the user may do with each (access). Kept in this browser (IndexedDB); every entry can be
 * downloaded as JSON and imported again.
 */
import { fabric, pbi } from "@/services/ms";
import { addSamples, readModel } from "@/services/model";
import { parseDatasetUrl, parseReportUrl, scope, type Ref } from "@/services/pbi";
import { definition, legacyDefinition } from "@/services/report-def";
import { describeReport } from "@/services/report-schema";
import { find, patch, put, type DatasetEntry, type Name, type ReportEntry } from "./store";
import type { ReportSchema } from "@/services/report-schema";

const message = (e: unknown) => (e instanceof Error ? e.message : String(e));

/* --------------------------------- Reports --------------------------------- */

/** What the user may do with a report: edit when its PBIR definition is readable (the editor needs it). */
async function reportAccess(ref: Ref): Promise<Pick<ReportEntry, "access" | "accessNote">> {
  if (ref.appId) return { access: "view", accessNote: "Report of a Power BI app: view only." };
  try {
    await definition(ref);
    return { access: "edit", accessNote: undefined };
  } catch (e) {
    // Readable in the older format: editable in Power BI, not by the editor's JSON tools.
    const legacy = await legacyDefinition(ref).then(() => true, () => false);
    return legacy
      ? { access: "view", accessNote: "Saved in the older PBIR-Legacy format: open it once in the Power BI editor and save it to convert it, then read it again." }
      : { access: "view", accessNote: `Definition not readable (${message(e).slice(0, 160)}): view only.` };
  }
}

/** Called with each report added and its schema (e.g. the Trainer makes an empty guide per visual). */
const reportListeners = new Set<(entry: ReportEntry, schema: ReportSchema) => unknown>();
export const onReportAdded = (listener: (entry: ReportEntry, schema: ReportSchema) => unknown) => (reportListeners.add(listener), () => void reportListeners.delete(listener));

export async function addReport(link: string): Promise<ReportEntry> {
  const ref = parseReportUrl(link);
  const r = await pbi(ref.appId ? `/apps/${ref.appId}/reports/${ref.id}` : `${scope(ref.groupId)}/reports/${ref.id}`);
  const old = await find("reports", ref.id);
  const entry = await put("reports", {
    ...old,
    id: ref.id,
    groupId: ref.groupId,
    appId: ref.appId,
    name: old?.renamed ? old.name : r.name,
    url: r.webUrl ?? link,
    datasetId: r.datasetId,
    datasetGroupId: r.datasetWorkspaceId ?? ref.groupId,
    contextIds: old?.contextIds ?? [],
    ...(await reportAccess(ref)),
  });
  // Its semantic model comes with it (in the background: reading a model takes a while).
  if (!(await find("datasets", r.datasetId))) void addDataset(r.datasetId, entry.datasetGroupId, `${r.name} (model)`).catch(() => undefined);
  // Its schema, then what the features do with a new report (e.g. a guide per visual), in the background too.
  void (old?.schema ? Promise.resolve(old.schema) : describeReport(entry)).then((schema) => Promise.all([...reportListeners].map((l) => l(entry, schema)))).catch(() => undefined);
  return entry;
}

/** Why a report cannot be edited by the Builder (a report of a Power BI app), with the way out. */
export const notEditable = (entry: ReportEntry) =>
  `"${entry.name}" is a report of a Power BI app: it cannot be edited. To build on the same data, create a new report from its semantic model (Builder › "New report from a semantic model…").`;

export async function refreshReport(entry: ReportEntry) {
  return addReport(entry.appId ? `https://app.powerbi.com/groups/me/apps/${entry.appId}/reports/${entry.id}` : `https://app.powerbi.com/groups/${entry.groupId ?? "me"}/reports/${entry.id}`);
}

/* -------------------------------- Datasets -------------------------------- */

/**
 * Adds a semantic model. Access: edit when its definition (TMDL) is readable, build when DAX queries
 * work (a model shared through an app), view otherwise (kept, with the reason). Value samples are
 * read in the background.
 */
export async function addDataset(linkOrId: string, groupId?: string, fallbackName?: string): Promise<DatasetEntry> {
  const ref = groupId ? { id: linkOrId, groupId } : parseDatasetUrl(linkOrId);
  const old = await find("datasets", ref.id);
  const d = await pbi(`${scope(ref.groupId)}/datasets/${ref.id}`).catch(() => null);
  const target: Ref = d ? ref : { id: ref.id };
  const base = { ...old, id: ref.id, groupId: target.groupId, name: old?.renamed ? old.name : (d?.name ?? old?.name ?? fallbackName ?? `Semantic model ${ref.id.slice(0, 8)}`), contextIds: old?.contextIds ?? [] };
  try {
    const info = await readModel(target, false);
    const entry = await put("datasets", { ...base, info, access: info.source === "tmdl" ? "edit" : "build", error: undefined });
    void addSamples(target, info).then(() => patch("datasets", ref.id, { info }), () => undefined);
    return entry;
  } catch (e) {
    return put("datasets", { ...base, access: "view", error: `Structure not readable: ${message(e)} (you need the Build permission on the model).` });
  }
}

export const refreshDataset = (entry: DatasetEntry) => addDataset(entry.id, entry.groupId, entry.name);

/* ------------------------------ Download / import ------------------------------ */

/** The collections whose entries can be downloaded and imported again (features add theirs). */
export type Kind = Name;
const KINDS = new Set<Kind>(["reports", "datasets"]);
export const registerKind = (kind: Kind) => void KINDS.add(kind);

/** Saves a file in the browser. */
export function downloadText(name: string, text: string, type = "application/json") {
  const url = URL.createObjectURL(new Blob([text], { type }));
  Object.assign(document.createElement("a"), { href: url, download: name.replace(/[^\p{L}\p{N}._ -]+/gu, "-") }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** One entry as a JSON file. */
export function downloadEntry(kind: Kind, entry: { id: string; name?: string; title?: string }) {
  downloadText(`${entry.name ?? entry.title ?? entry.id}.${kind.slice(0, -1)}.json`, JSON.stringify({ hub: kind, entry }, null, 2));
}

/** Entries downloaded before (one per file, or a whole collection), put back in the library. */
export async function importEntries(files: FileList | File[]) {
  let count = 0;
  for (const file of Array.from(files)) {
    const data = JSON.parse(await file.text());
    const items: { hub: Kind; entry: { id: string } }[] = Array.isArray(data) ? data : [data];
    for (const item of items) {
      if (!KINDS.has(item?.hub) || typeof item.entry?.id !== "string") throw new Error(`${file.name}: not an entry downloaded from the library.`);
      await put(item.hub as Name, { contextIds: [], ...item.entry } as never);
      count++;
    }
  }
  return count;
}
