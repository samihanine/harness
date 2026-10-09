/** Library entries: what is read from Power BI / SharePoint when a link is added or refreshed. */
import { readTable, resolveExcel, saveFields } from "./excel";
import { COLORS, type Field } from "./schema";
import { graph, pbi } from "./ms";
import { readModel, modelSummary } from "./model";
import { embedReport, parseDatasetUrl, parseReportUrl, reportText, resetEmbed, scope, type Ref } from "./pbi";
import { definition, definitionText } from "./report-def";
import { find, patch, put, type DatasetEntry, type ExcelEntry, type ReportEntry } from "./store";

/**
 * What is stored to rebuild a report: its full definition (formatting, shapes, text, theme,
 * interactions…) when it can be read (edit rights), otherwise the embed summary (Viewer rights:
 * pages, visual types, positions, fields, filters only).
 */
async function snapshot(ref: Ref) {
  if (ref.appId) return `Report of a Power BI app: its definition is not readable (only the workspace report is). Summary only (no formatting, text or theme):\n${await embedSnapshot(ref)}`;
  try {
    return definitionText(await definition(ref));
  } catch (e) {
    const why = e instanceof Error ? e.message : String(e);
    return `Full definition not readable (${why}). Summary only (no formatting, text or theme):\n${await embedSnapshot(ref)}`;
  }
}

/** Reads a report through a hidden embed: works with Viewer / Build rights. */
async function embedSnapshot(ref: Ref) {
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
  const r = await pbi(ref.appId ? `/apps/${ref.appId}/reports/${ref.id}` : `${scope(ref.groupId)}/reports/${ref.id}`);
  const old = await find("reports", ref.id);
  const entry: ReportEntry = {
    ...old,
    id: ref.id,
    groupId: ref.groupId,
    appId: ref.appId,
    name: r.name,
    url: r.webUrl ?? link,
    datasetId: r.datasetId,
    datasetGroupId: r.datasetWorkspaceId ?? ref.groupId,
    editable: ref.appId ? false : (r.isOwnedByMe ?? undefined),
    context: old?.context ?? "",
    snapshot: await snapshot(ref).catch((e) => `Not readable: ${e instanceof Error ? e.message : e}`),
  };
  await put("reports", entry);
  if (!(await find("datasets", r.datasetId))) {
    // The report is kept even when its model cannot be read; the reason is shown in the library.
    const error = await addDataset(r.datasetId, entry.datasetGroupId, {}, `${r.name} (model)`).then(
      (d) => (d.groupId !== entry.datasetGroupId ? patch("reports", entry.id, { datasetGroupId: d.groupId }) : undefined, undefined),
      (e) => (e instanceof Error ? e.message : String(e)),
    );
    await patch("reports", entry.id, { datasetError: error });
  }
  return entry;
}

/**
 * Adds a semantic model. Without access to its workspace (a model shared through an app, Build right
 * only), it is read with DAX queries, which only need the Build right; `fallbackName` names it then.
 */
export async function addDataset(linkOrId: string, groupId?: string, values: Partial<DatasetEntry> = {}, fallbackName?: string): Promise<DatasetEntry> {
  const ref = groupId ? { id: linkOrId, groupId } : parseDatasetUrl(linkOrId);
  const old = await find("datasets", ref.id);
  const d = await pbi(`${scope(ref.groupId)}/datasets/${ref.id}`).catch(() => null);
  const target: Ref = d ? ref : { id: ref.id };
  const info = await readModel(target).catch((e) => {
    throw new Error(`Semantic model not readable: no access to its workspace, and DAX queries fail (${e instanceof Error ? e.message : e}). You need the Build permission on it.`);
  });
  return put("datasets", {
    ...old,
    id: ref.id,
    groupId: target.groupId,
    name: d?.name ?? old?.name ?? fallbackName ?? `Semantic model ${ref.id.slice(0, 8)}`,
    context: old?.context ?? "",
    ...values,
    info,
    model: modelSummary(info),
  });
}

export async function refreshDatasetText(entry: DatasetEntry) {
  const info = await readModel(entry);
  await patch("datasets", entry.id, { info, model: modelSummary(info) });
}

/** Adds an Excel table: its columns are detected, written as its schema and the sheet is styled after them. */
/**
 * Adds an Excel table: its columns are detected, written as its schema and the sheet is styled after them.
 * `required`: fields the file must have (missing columns are added; a file without table gets one).
 */
export async function addExcel(link: string, table?: string, required: Field[] = []): Promise<ExcelEntry> {
  let file = await resolveExcel(link);
  if (file.tables.length === 0 && required.length) {
    await createTable(file, required.map((f) => f.name));
    file = await resolveExcel(link);
  }
  if (file.tables.length === 0) throw new Error("This file has no Excel table: select the data in Excel and use Insert › Table.");
  const entry = await put("excels", {
    id: crypto.randomUUID(),
    name: file.name.replace(/\.xlsx$/i, ""),
    link,
    driveId: file.driveId,
    itemId: file.itemId,
    fileUrl: file.fileUrl,
    table: table && file.tables.includes(table) ? table : file.tables[0],
    context: "",
  });
  const data = await readTable(entry);
  const missing = required.filter((r) => !data.fields.some((f) => f.name.toLowerCase() === r.name.toLowerCase()));
  if (!data.hasSchema || missing.length) await saveFields(entry, data, [...data.fields, ...missing]).catch(() => undefined);
  return entry;
}

/** A table on the first sheet of an empty workbook, with the given header. */
async function createTable(file: { driveId: string; itemId: string }, header: string[]) {
  const book = `/drives/${file.driveId}/items/${file.itemId}/workbook`;
  const { value } = await graph(`${book}/worksheets?$select=name`);
  const sheet = `${book}/worksheets('${encodeURIComponent(value[0].name)}')`;
  const address = `A1:${String.fromCharCode(64 + header.length)}1`;
  await graph(`${sheet}/range(address='${address}')`, "PATCH", { values: [header] });
  await graph(`${sheet}/tables/add`, "POST", { address, hasHeaders: true });
}

/** Fields of the viewer's Excel files. */
export const INFO_FIELDS: Field[] = [
  { name: "label", type: "string", required: true },
  { name: "url", type: "string", required: true },
  { name: "icon", type: "string", description: "Lucide icon name (e.g. book-open, mail, video)" },
  { name: "color", type: "option", options: COLORS.map((c) => ({ value: c, color: c })) },
  { name: "description", type: "string" },
];
export const GUIDE_FIELDS: Field[] = [
  { name: "title", type: "string", required: true },
  { name: "content", type: "text" },
  { name: "links", type: "string", description: "pageName or pageName/visualName, separated by ;" },
];

/** Gives a report its info links or guides Excel file (from a link, or one of the library). */
export async function connectReportExcel(reportId: string, kind: "info" | "guides", linkOrId: string) {
  const known = await find("excels", linkOrId);
  const fields = kind === "info" ? INFO_FIELDS : GUIDE_FIELDS;
  const excel = known ? await ensureFields(known, fields) : await addExcel(linkOrId, undefined, fields);
  await patch("reports", reportId, kind === "info" ? { infoExcelId: excel.id } : { guidesExcelId: excel.id });
}

async function ensureFields(excel: ExcelEntry, required: Field[]) {
  const data = await readTable(excel);
  const missing = required.filter((r) => !data.fields.some((f) => f.name.toLowerCase() === r.name.toLowerCase()));
  if (missing.length) await saveFields(excel, data, [...data.fields, ...missing]);
  return excel;
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
