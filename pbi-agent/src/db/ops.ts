/** Adding and refreshing reports and datasets in the local database. */
import type { ReportRef } from "@/pbi/api";
import { getDataset, getReport, parseReportUrl } from "@/pbi/api";
import { extractReport } from "@/pbi/extract";
import { readModel } from "@/pbi/model";
import { measuresOf, readPbix } from "@/pbi/pbix";
import type { Connection, ReportContent } from "@/pbi/types";
import { importExcel, saveDatasetFile } from "@/query/local";
import { db } from "./db";

type Log = (step: string) => void;
const now = () => new Date().toISOString();

/** Stores a report and makes sure its dataset is in the database too. */
async function saveReport(
  { reportId, title, workspaceId, content, connection }: { reportId: string; title: string; workspaceId?: string; content: ReportContent; connection?: Connection },
  log: Log,
) {
  const datasetId = connection?.kind === "remote" ? connection.datasetId : undefined;
  db.upsert("reports", reportId, {
    title,
    dataset_id: datasetId ?? "",
    workspace_id: workspaceId ?? "",
    config_json: await db.writeJson("reports", reportId, "config", { reportId, workspaceId, ...connection }),
    content_json: await db.writeJson("reports", reportId, "content", content),
    updated_at: now(),
  });
  if (datasetId && !db.row("datasets", datasetId)) {
    log("Adding its dataset…");
    await saveRemoteDataset({ datasetId, workspaceId: connection?.kind === "remote" ? connection.workspaceId : undefined, knownMeasures: measuresOf(content), fallbackTitle: `Model of ${title}` }, log);
  }
  return reportId;
}

/** Report from a .pbix file connected live to a published model. */
export async function addReportFromPbix(file: File, log: Log = () => {}) {
  log("Reading the file…");
  const info = readPbix(new Uint8Array(await file.arrayBuffer()));
  if (info.connection?.kind !== "remote") throw new Error(info.content.warnings.join(" ") || "This report is not connected to a published semantic model.");
  const reportId = info.reportId ?? crypto.randomUUID();
  // The published report (same id) gives the workspace when it is accessible.
  const published = info.reportId ? await getReport({ reportId: info.reportId }).catch(() => null) : null;
  return saveReport({ reportId, title: published?.name ?? file.name.replace(/\.pbix$/i, ""), content: info.content, connection: info.connection }, log);
}

/** Report from its link or id: full definition when downloadable, else read through embedding. */
export async function addReportFromUrl(url: string, log: Log = () => {}) {
  const { reportId, groupId } = parseReportUrl(url);
  return refreshReport({ reportId, groupId }, log);
}

export async function refreshReport(ref: ReportRef, log: Log = () => {}) {
  log("Reading the report…");
  const meta = await getReport(ref).catch(() => null);
  const { content, connection } = await extractReport(ref, log);
  const datasetId = (connection?.kind === "remote" ? connection.datasetId : undefined) ?? meta?.datasetId;
  if (!datasetId) content.warnings.push("Dataset unknown: the report metadata is not readable with your rights.");
  const existing = db.row("reports", ref.reportId);
  return saveReport(
    {
      reportId: ref.reportId,
      title: (existing?.title as string) || meta?.name || "Report",
      workspaceId: ref.groupId,
      content,
      connection: datasetId ? { kind: "remote", datasetId, workspaceId: meta?.datasetWorkspaceId ?? ref.groupId } : undefined,
    },
    log,
  );
}

async function saveRemoteDataset(
  { datasetId, workspaceId, knownMeasures = [], fallbackTitle }: { datasetId: string; workspaceId?: string; knownMeasures?: string[]; fallbackTitle?: string },
  log: Log,
) {
  const ref = { datasetId, groupId: workspaceId };
  const meta = await getDataset(ref).catch(() => null);
  log("Reading the semantic model…");
  const model = await readModel(ref, knownMeasures);
  const existing = db.row("datasets", datasetId);
  db.upsert("datasets", datasetId, {
    title: (existing?.title as string) || meta?.name || fallbackTitle || "Dataset",
    source: "remote",
    workspace_id: workspaceId ?? "",
    config_json: await db.writeJson("datasets", datasetId, "config", { kind: "remote", datasetId, workspaceId, datasetName: meta?.name } satisfies Connection),
    content_json: await db.writeJson("datasets", datasetId, "content", model),
    updated_at: now(),
  });
  return datasetId;
}

export async function refreshDataset(id: string, log: Log = () => {}) {
  const row = db.row("datasets", id);
  if (!row) throw new Error("Unknown dataset");
  if (row.source === "excel") throw new Error("Re-import the Excel file to refresh a local dataset.");
  const reports = db.rows("reports").filter((r) => r.dataset_id === id);
  const known = (await Promise.all(reports.map((r) => db.readJson<ReportContent>(r.content_json)))).flatMap((c) => (c ? measuresOf(c) : []));
  return saveRemoteDataset({ datasetId: id, workspaceId: (row.workspace_id as string) || undefined, knownMeasures: known }, log);
}

/** Local dataset from an Excel file (sheets become tables; converted to Excel tables when needed). */
export async function addExcelDataset(handle: FileSystemFileHandle, log: Log = () => {}) {
  log("Reading the workbook…");
  const id = `excel-${crypto.randomUUID().slice(0, 8)}`;
  const { model, converted, sheets } = await importExcel(handle);
  if (converted) model.warnings.push(`${converted} sheet(s) were converted to Excel tables in ${handle.name}.`);
  await saveDatasetFile(id, handle);
  db.upsert("datasets", id, {
    title: handle.name.replace(/\.xlsx$/i, ""),
    source: "excel",
    config_json: await db.writeJson("datasets", id, "config", { kind: "excel", fileName: handle.name, sheets } satisfies Connection),
    content_json: await db.writeJson("datasets", id, "content", model),
    updated_at: now(),
  });
  return id;
}
