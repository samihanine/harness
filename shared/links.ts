/**
 * Power BI links → ids: the browser's address of a report or a model (a workspace's, an app's, "My workspace",
 * shared), an embed link (reportEmbed?reportId=…&groupId=…), or a bare id. Used by the app and the server.
 */
const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

export type Link = { report?: string; dataset?: string; bare?: string; appId?: string; groupId?: string; visual?: string; page?: string };

export function parseLink(url: string): Link | undefined {
  const g = (re: RegExp) => url.match(re)?.[1]?.toLowerCase();
  const report = g(/(?:reports|rdlreports)\/([0-9a-f-]{36})/i) ?? g(/[?&]reportId=([0-9a-f-]{36})/i);
  const dataset = g(/(?:datasets|semanticmodels)\/([0-9a-f-]{36})/i) ?? g(/[?&]datasetId=([0-9a-f-]{36})/i);
  const appId = g(/apps\/([0-9a-f-]{36})/i) ?? g(/[?&]appId=([0-9a-f-]{36})/i);
  const groupId = g(/groups\/([0-9a-f-]{36})/i) ?? g(/[?&]groupId=([0-9a-f-]{36})/i);
  const bare = !report && !dataset && !appId && !groupId ? url.match(GUID)?.[0]?.toLowerCase() : undefined;
  // A report page and a visual: …/reports/<id>/<page>[/<visual>]
  const rest = url.match(/reports\/[0-9a-f-]{36}\/([^/?#]+)(?:\/([^/?#]+))?/i);
  return report || dataset || bare ? { report, dataset, bare, appId, groupId, ...(rest ? { page: rest[1], visual: rest[2] } : {}) } : undefined;
}

export const isPowerBiLink = (text: string) => /powerbi\.com/i.test(text) || GUID.test(text);

/** The address of a report or a model in Power BI, from its ids. */
export const reportUrl = (r: { id: string; groupId?: string; appId?: string }) => (r.appId ? `https://app.powerbi.com/groups/me/apps/${r.appId}/reports/${r.id}` : `https://app.powerbi.com/groups/${r.groupId ?? "me"}/reports/${r.id}`);
export const datasetUrl = (d: { id: string; groupId?: string }) => `https://app.powerbi.com/groups/${d.groupId ?? "me"}/datasets/${d.id}/details`;
