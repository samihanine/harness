/** Power BI REST: references, metadata, DAX queries (executeQueries). */
import { getToken } from "./auth";

export const PBI = "https://api.powerbi.com/v1.0/myorg";
const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

export type DatasetRef = { datasetId: string; groupId?: string };
export type ReportRef = { reportId: string; groupId?: string };

const groupOf = (url: string) => {
  const id = url.match(/groups\/([^/?#]+)/i)?.[1] ?? url.match(/[?&]groupId=([^&#]+)/i)?.[1];
  return id && GUID.test(id) ? id.toLowerCase() : undefined;
};

/** Report link (app.powerbi.com/groups/…/reports/<id>/…, reportEmbed?reportId=…) or bare id. */
export function parseReportUrl(url: string) {
  const reportId = url.match(/reports\/([0-9a-f-]{36})/i)?.[1] ?? url.match(/reportId=([0-9a-f-]{36})/i)?.[1] ?? url.match(GUID)?.[0];
  if (!reportId) throw new Error("No report id found in this link");
  return {
    reportId: reportId.toLowerCase(),
    groupId: groupOf(url),
    pageName: url.match(/reports\/[0-9a-f-]{36}\/([^/?#]+)/i)?.[1] ?? url.match(/[?&]pageName=([^&#]+)/i)?.[1],
    visualId: url.match(/[?&]visual=([^&#]+)/i)?.[1],
  };
}

export const embedUrl = (ref: ReportRef) =>
  `https://app.powerbi.com/reportEmbed?reportId=${ref.reportId}${ref.groupId ? `&groupId=${ref.groupId}` : ""}`;
export const scope = (groupId?: string) => (groupId ? `${PBI}/groups/${groupId}` : PBI);

export async function pbiFetch<T>(url: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch(url, {
    ...init,
    headers: { Authorization: `Bearer ${await getToken()}`, "Content-Type": "application/json", ...init.headers },
  });
  const text = await response.text();
  let data: Record<string, any> = {};
  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = { message: text.slice(0, 300) };
  }
  if (!response.ok)
    throw new Error(
      data.error?.pbi?.error?.details?.[0]?.detail?.value ?? data.error?.message ?? data.message ?? `Power BI error ${response.status}`,
    );
  return data as T;
}

export const getReport = (ref: ReportRef) =>
  pbiFetch<{ id: string; name: string; datasetId: string; datasetWorkspaceId?: string; webUrl?: string }>(
    `${scope(ref.groupId)}/reports/${ref.reportId}`,
  );

export const getDataset = (ref: DatasetRef) =>
  pbiFetch<{ id: string; name: string; configuredBy?: string; webUrl?: string }>(`${scope(ref.groupId)}/datasets/${ref.datasetId}`);

/** Downloads the .pbix of a report (needs download rights); through the local API (CORS). */
export async function exportReport(ref: ReportRef) {
  const path = `/v1.0/myorg${ref.groupId ? `/groups/${ref.groupId}` : ""}/reports/${ref.reportId}/Export?downloadType=LiveConnect`;
  const response = await fetch(`/api/pbi${path}`);
  if (!response.ok) throw new Error(`Export not allowed (${response.status})`);
  return new Uint8Array(await response.arrayBuffer());
}

export type QueryResult = { columns: string[]; rows: Record<string, unknown>[] };

/** "Table[Column]" / "[Measure]" result keys → readable names. */
const clean = (key: string) => key.replace(/^.*\[(.*)\]$/, "$1");

/** DAX query (EVALUATE …) with executeQueries: Build or Read permission is enough. */
export async function executeDax(ref: DatasetRef, dax: string): Promise<QueryResult> {
  const data = await pbiFetch<{ results: { tables?: { rows: Record<string, unknown>[] }[]; error?: { message: string } }[] }>(
    `${scope(ref.groupId)}/datasets/${ref.datasetId}/executeQueries`,
    { method: "POST", body: JSON.stringify({ queries: [{ query: dax }], serializerSettings: { includeNulls: true } }) },
  );
  const result = data.results[0];
  if (result?.error) throw new Error(result.error.message);
  const raw = result?.tables?.[0]?.rows ?? [];
  const keys = raw[0] ? Object.keys(raw[0]) : [];
  // Keep unique readable names (two columns may share a name in different tables).
  const names = keys.map((key, i) => (keys.findIndex((k) => clean(k) === clean(key)) === i ? clean(key) : key));
  return { columns: names, rows: raw.map((row) => Object.fromEntries(keys.map((k, i) => [names[i], row[k]]))) };
}
