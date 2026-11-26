/**
 * The Builder works on a .pbix the user drops (a report live-connected to a semantic model, or with its own
 * model): it is imported privately into the user's "My workspace" (Imports API, works on Pro) so the app shows it
 * embedded, as Power BI draws it; the Builder's tools then change that copy's definition (PBIR, or the old
 * format shown as PBIR), and the .pbix to download is the dropped file with the copy's current definition put
 * back in — its connection to the model untouched. The model itself is never changed: a Build right on it is
 * enough (report measures live in the report).
 */
import JSZip from "jszip";
import type { Ref } from "@shared/types";
import { pbi, request } from "../ms/api";
import { anyReportDefinition, partText, type Part } from "./definitions";

export type PbixInfo = { format: "PBIR" | "PBIR-Legacy"; datasetId?: string; live: boolean; pages: number };

/** Legacy pbix files keep their Layout in UTF-16 LE. */
const utf16 = (bytes: Uint8Array) => new TextDecoder("utf-16le").decode(bytes).replace(/^﻿/, "");
const toUtf16 = (text: string) => {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = c >> 8;
  }
  return out;
};

/** What a .pbix holds: its report format, the model it uses (live connection), its pages. */
export async function pbixInfo(data: Uint8Array): Promise<PbixInfo> {
  const zip = await JSZip.loadAsync(data).catch(() => {
    throw new Error("This is not a Power BI file (.pbix).");
  });
  const layout = zip.file("Report/Layout");
  const pbir = Object.keys(zip.files).some((f) => f.startsWith("Report/definition/"));
  if (!layout && !pbir) throw new Error("This .pbix holds no report (a .pbit template, or a damaged file?).");
  const connections = zip.file("Connections") ? JSON.parse((await zip.file("Connections")!.async("string")).replace(/^﻿/, "")) : {};
  const remote = connections.RemoteArtifacts?.[0]?.DatasetId as string | undefined;
  const live = connections.Connections?.find((c: { PbiModelDatabaseName?: string; ConnectionString?: string }) => c.PbiModelDatabaseName || /pbiazure/i.test(c.ConnectionString ?? ""));
  const datasetId = (live?.PbiModelDatabaseName as string | undefined) ?? live?.ConnectionString?.match(/Initial Catalog=(?:sobe_wowvirtualserver-)?([0-9a-f-]{36})/i)?.[1] ?? remote;
  const pages = layout ? (JSON.parse(utf16(await layout.async("uint8array"))).sections?.length ?? 0) : Object.keys(zip.files).filter((f) => /^Report\/definition\/pages\/[^/]+\/page\.json$/.test(f)).length;
  return { format: layout ? "PBIR-Legacy" : "PBIR", datasetId, live: !!live && !zip.file("DataModel"), pages };
}

/**
 * Imports a .pbix into My workspace under a name of its own (replaced when imported again): the report the app
 * embeds and the Builder changes. Returns its id and its model's.
 */
export async function importPbix(data: Uint8Array, name: string): Promise<Ref & { datasetId?: string; datasetGroupId?: string }> {
  const form = new FormData();
  form.append("file", new Blob([data as unknown as BlobPart]), `${name}.pbix`);
  const r = await request("powerbi", `/imports?datasetDisplayName=${encodeURIComponent(`${name}.pbix`)}&nameConflict=CreateOrOverwrite&skipReport=false`, { method: "POST", body: form });
  const body = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(`Power BI refused the file (${r.status}): ${body?.error?.message ?? body?.error?.code ?? JSON.stringify(body).slice(0, 300)}`);
  for (let i = 0; i < 120; i++) {
    await new Promise((res) => setTimeout(res, i < 10 ? 1000 : 2000));
    const state = await pbi(`/imports/${body.id}`);
    if (state.importState === "Failed") throw new Error(`Power BI could not import the file: ${state.error?.details ?? state.error?.code ?? "import failed"}. A report live-connected to a model needs the Build permission on it.`);
    if (state.importState === "Succeeded") {
      const report = state.reports?.[0];
      if (!report) throw new Error("The file was imported but holds no report.");
      const meta = await pbi(`/reports/${report.id}`);
      return { id: report.id, datasetId: meta.datasetId, datasetGroupId: meta.datasetWorkspaceId ?? undefined };
    }
  }
  throw new Error("The import into Power BI did not finish in 4 minutes.");
}

/** Removes the private copy (the item was removed from the Builder). */
export async function forgetPbixCopy(id: string) {
  await request("powerbi", `/reports/${id}`, { method: "DELETE" }).catch(() => undefined);
}

/**
 * The dropped .pbix with the copy's current definition in it: the report's files replaced (old format: the
 * Layout; PBIR: Report/definition), the security bindings removed (Power BI Desktop refuses a changed file
 * that keeps them); the model and its connection as they were.
 */
export async function pbixWithDefinition(original: Uint8Array, copy: Ref): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(original);
  const parts: Part[] = await anyReportDefinition(copy);
  const report = parts.find((p) => p.path === "report.json");
  const remove = (prefix: string) => Object.keys(zip.files).filter((f) => f === prefix || f.startsWith(`${prefix}/`)).forEach((f) => zip.remove(f));
  if (report) {
    remove("Report/definition");
    zip.file("Report/Layout", toUtf16(partText(report)));
  } else {
    remove("Report/definition");
    zip.remove("Report/Layout");
    for (const p of parts.filter((x) => x.path.startsWith("definition/"))) zip.file(`Report/${p.path}`, Buffer.from(p.payload, "base64"));
  }
  for (const p of parts.filter((x) => x.path.startsWith("StaticResources/"))) zip.file(`Report/${p.path}`, Buffer.from(p.payload, "base64"));
  zip.remove("SecurityBindings");
  const types = zip.file("[Content_Types].xml");
  if (types) {
    let xml = (await types.async("string")).replace(/<Override PartName="\/SecurityBindings"[^>]*\/>/, "");
    if (report && !/PartName="\/Report\/Layout"/.test(xml)) xml = xml.replace("</Types>", '<Override PartName="/Report/Layout" ContentType="" /></Types>');
    if (!report) xml = xml.replace(/<Override PartName="\/Report\/Layout"[^>]*\/>/, "");
    if (!/Extension="json"/.test(xml)) xml = xml.replace(/(<Types[^>]*>)/, '$1<Default Extension="json" ContentType="" />');
    zip.file("[Content_Types].xml", xml);
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
