/**
 * The Data Scientist's workbook: an .xlsx in the user's OneDrive ("Hub workbooks"), changed through Excel's own
 * API (Graph workbook): formulas, formats, conditional formats, validations, tables, charts — and computed by
 * Excel itself, also its CUBE formulas on the Power BI models (with the user's identity) and formulas over the
 * hundreds of thousands of rows of an uploaded file. Connections to the models are put in the file when a model
 * is added (Excel's API cannot add them). The file downloaded is the real workbook, connected and refreshable.
 */
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { accessToken } from "../ms/auth";
import { graph, request } from "../ms/api";

const FOLDER = "Hub workbooks";
export type WorkbookRef = { itemId: string; webUrl?: string };

/* --------------------------------- OneDrive --------------------------------- */

async function put(path: string, data: Uint8Array): Promise<{ id: string; webUrl: string }> {
  for (let attempt = 0; ; attempt++) {
    const r = await request("graph", `/me/drive/root:/${encodeURIComponent(FOLDER)}/${encodeURIComponent(path)}:/content?@microsoft.graph.conflictBehavior=replace`, { method: "PUT", headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }, body: data as unknown as BodyInit });
    const body = await r.json().catch(() => ({}));
    if (r.ok) return body;
    // Busy, or the file still open in a session: a moment later.
    if (attempt < 4 && [409, 423, 429, 503].includes(r.status)) {
      await new Promise((res) => setTimeout(res, 1500 * (attempt + 1)));
      continue;
    }
    throw new Error(`OneDrive refused the workbook (${r.status}): ${body?.error?.message ?? ""}`);
  }
}

/** The workbook's file as it is now. */
export async function workbookBytes(itemId: string) {
  const r = await request("graph", `/me/drive/items/${itemId}/content`);
  if (!r.ok) throw new Error(`The workbook could not be read from OneDrive (${r.status}).`);
  return new Uint8Array(await r.arrayBuffer());
}

/** A new workbook for an analysis (one sheet), in OneDrive. */
export async function createWorkbook(key: string, title: string): Promise<WorkbookRef> {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet("Analysis");
  const item = await put(`${key}.xlsx`, new Uint8Array(await wb.xlsx.writeBuffer()));
  return { itemId: item.id, webUrl: item.webUrl, ...(title ? {} : {}) };
}

/** The workbook replaced by new bytes (a change made on the file itself: connections, sheets of a file). */
export async function replaceWorkbook(itemId: string, data: Uint8Array) {
  const meta = await graph(`/me/drive/items/${itemId}?$select=name`);
  await put(meta.name, data);
}

/* -------------------------------- Connections -------------------------------- */

/** The tenant of the user (Power BI's connection names it). */
async function tenant() {
  try {
    return JSON.parse(Buffer.from((await accessToken("powerbi")).split(".")[1], "base64url").toString()).tid as string;
  } catch {
    return "common";
  }
}

/** A model's connection name in the workbook (CUBE formulas name it): "PBI " + the model's name. */
export const connectionName = (model: { name: string }) => `PBI ${model.name.replace(/["<>&]/g, "").slice(0, 60)}`;

/** Power BI's connection for Excel ("Analyze in Excel": works on Pro). */
const connectionString = (datasetId: string, tid: string) =>
  [
    "Provider=MSOLAP.8",
    "Integrated Security=ClaimsToken",
    "Persist Security Info=True",
    `Initial Catalog=sobe_wowvirtualserver-${datasetId}`,
    "Data Source=pbiazure://api.powerbi.com",
    "MDX Compatibility=1",
    "Safety Options=2",
    "MDX Missing Member Mode=Error",
    `Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, ${tid}`,
    "Update Isolation Level=2",
  ].join(";");

const xml = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

/** Makes sure the workbook has a connection to each model (added to the file, then the file put back). */
export async function ensureConnections(itemId: string, models: { id: string; name: string }[]) {
  if (!models.length) return [];
  const zip = await JSZip.loadAsync(await workbookBytes(itemId));
  const current = (await zip.file("xl/connections.xml")?.async("string")) ?? "";
  const missing = models.filter((m) => !current.includes(`sobe_wowvirtualserver-${m.id}`));
  if (!missing.length) return models.map(connectionName);
  const tid = await tenant();
  let id = Math.max(0, ...[...current.matchAll(/<connection [^>]*id="(\d+)"/g)].map((m) => Number(m[1])));
  const added = missing.map((m) => `<connection id="${++id}" keepAlive="1" name="${xml(connectionName(m))}" type="5" refreshedVersion="8" background="1"><dbPr connection="${xml(connectionString(m.id, tid))}" command="Model" commandType="1"/><olapPr sendLocale="1" rowDrillCount="1000"/></connection>`).join("");
  if (current) zip.file("xl/connections.xml", current.replace("</connections>", `${added}</connections>`));
  else {
    zip.file("xl/connections.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${added}</connections>`);
    const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
    zip.file("xl/_rels/workbook.xml.rels", rels.replace("</Relationships>", `<Relationship Id="rIdHubConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/></Relationships>`));
    const types = await zip.file("[Content_Types].xml")!.async("string");
    zip.file("[Content_Types].xml", types.replace("</Types>", `<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/></Types>`));
  }
  await closeSession(itemId);
  await replaceWorkbook(itemId, await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
  return models.map(connectionName);
}

/* ------------------------------ Files as sheets ------------------------------ */

/** Sheets of an uploaded Excel / CSV file copied into the workbook (values, formulas, formats, widths). */
export async function addFileSheets(itemId: string, file: { name: string; data: Uint8Array }) {
  const target = new ExcelJS.Workbook();
  await target.xlsx.load((await workbookBytes(itemId)) as never);
  const source = new ExcelJS.Workbook();
  if (/\.(csv|tsv|txt)$/i.test(file.name)) {
    const ws = source.addWorksheet(file.name.replace(/\.[^.]+$/, ""));
    const text = new TextDecoder().decode(file.data);
    const sep = /\.tsv$/i.test(file.name) || text.split("\n")[0].split("\t").length > text.split("\n")[0].split(";").length ? "\t" : text.split("\n")[0].split(";").length > text.split("\n")[0].split(",").length ? ";" : ",";
    for (const line of text.split(/\r?\n/)) if (line) ws.addRow(line.split(sep).map((v) => (v !== "" && !Number.isNaN(Number(v)) ? Number(v) : v)));
  } else await source.xlsx.load(file.data as never);
  const names: string[] = [];
  source.eachSheet((s) => {
    let name = s.name.slice(0, 31);
    for (let n = 2; target.getWorksheet(name); n++) name = `${s.name.slice(0, 27)} ${n}`;
    const t = target.addWorksheet(name);
    s.columns?.forEach((c, i) => c.width && (t.getColumn(i + 1).width = c.width));
    s.eachRow({ includeEmpty: false }, (row, r) => {
      const out = t.getRow(r);
      if (row.height) out.height = row.height;
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        const o = out.getCell(c);
        o.value = cell.value;
        o.style = cell.style;
      });
    });
    names.push(name);
  });
  await closeSession(itemId);
  await replaceWorkbook(itemId, new Uint8Array(await target.xlsx.writeBuffer()));
  return names;
}

/* ------------------------------ Excel's own API ------------------------------ */

/** One session per workbook (changes kept), reused by the calls of a turn and closed after a while. */
const sessions = new Map<string, { id: string; at: number }>();
async function session(itemId: string) {
  const s = sessions.get(itemId);
  if (s && Date.now() - s.at < 4 * 60_000) return (s.at = Date.now()), s.id;
  const r = await graph(`/me/drive/items/${itemId}/workbook/createSession`, "POST", { persistChanges: true });
  sessions.set(itemId, { id: r.id, at: Date.now() });
  return r.id as string;
}
export async function closeSession(itemId: string) {
  const s = sessions.get(itemId);
  sessions.delete(itemId);
  if (s) await request("graph", `/me/drive/items/${itemId}/workbook/closeSession`, { method: "POST", headers: { "workbook-session-id": s.id } }).catch(() => undefined);
}

/** A call of Excel's API on the workbook, in its session (a lost session is opened again once). */
export async function excel<T = any>(itemId: string, path: string, method = "GET", body?: unknown): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    const sid = await session(itemId);
    const r = await request("graph", `/me/drive/items/${itemId}/workbook${path}`, { method, headers: { "workbook-session-id": sid, ...(body !== undefined ? { "Content-Type": "application/json" } : {}) }, ...(body !== undefined ? { body: JSON.stringify(body) } : {}) });
    const text = await r.text();
    const json = text ? JSON.parse(text) : null;
    if (r.ok) return json as T;
    const message = json?.error?.innerError?.message ?? json?.error?.message ?? text.slice(0, 300);
    if (attempt === 0 && (r.status === 404 && /session/i.test(message) || /SessionNotFound|InvalidSession/i.test(String(json?.error?.code)))) {
      sessions.delete(itemId);
      continue;
    }
    if (attempt < 3 && [429, 503, 504].includes(r.status)) {
      await new Promise((res) => setTimeout(res, (Number(r.headers.get("Retry-After")) || 2) * 1000));
      continue;
    }
    throw new Error(`Excel: ${message}`);
  }
}

export const sheetPath = (sheet: string) => `/worksheets('${encodeURIComponent(sheet.replace(/'/g, "''"))}')`;
export const rangePath = (sheet: string, address: string) => `${sheetPath(sheet)}/range(address='${encodeURIComponent(address.replace(/'/g, "''"))}')`;
