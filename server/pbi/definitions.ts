/** Definitions through the Fabric API: semantic models as TMDL, reports as PBIR (read and written whole). */
import type { Ref } from "@shared/types";
import { fabric, workspaceId } from "../ms/api";

export type ModelFile = { path: string; text: string };
export type Part = { path: string; payload: string; payloadType: string };

export const decode = (payload: string) => Buffer.from(payload, "base64").toString("utf8");
export const encode = (text: string) => Buffer.from(text, "utf8").toString("base64");
export const partText = (p: Part) => decode(p.payload);
export const textPart = (path: string, text: string): Part => ({ path, payload: encode(text), payloadType: "InlineBase64" });

export async function readModelFiles(datasetId: string, groupId?: string): Promise<ModelFile[]> {
  const result = await fabric(`/workspaces/${await workspaceId(groupId)}/semanticModels/${datasetId}/getDefinition?format=TMDL`, "POST");
  return result.definition.parts.filter((p: Part) => p.path !== ".platform").map((p: Part) => ({ path: p.path, text: decode(p.payload) }));
}

export async function writeModelFiles(datasetId: string, files: ModelFile[], groupId?: string) {
  await fabric(`/workspaces/${await workspaceId(groupId)}/semanticModels/${datasetId}/updateDefinition`, "POST", { definition: { parts: files.map((f) => textPart(f.path, f.text)) } }).catch((e) => {
    const m = e instanceof Error ? e.message : String(e);
    if (/content provider or type not supported/i.test(m)) throw new Error("Power BI does not let this model be changed through its API (a sample model, or a kind of model it cannot write). Give the DAX to the user to add it in Power BI Desktop.");
    throw e;
  });
}

/** A report's definition in PBIR format (refused for reports not editable, or with an encrypted sensitivity label). */
export async function reportDefinition(report: Ref): Promise<Part[]> {
  const ws = await workspaceId(report.groupId);
  const result = await fabric(`/workspaces/${ws}/reports/${report.id}/getDefinition?format=PBIR`, "POST").catch((e) => {
    throw new Error(`The report definition is not readable in PBIR format (${e instanceof Error ? e.message : e}). You need to be able to edit the report; a report saved in the old format must be saved once from Power BI.`);
  });
  return result.definition.parts.filter((p: Part) => p.path !== ".platform");
}

export async function writeReportDefinition(report: Ref, parts: Part[]) {
  await fabric(`/workspaces/${await workspaceId(report.groupId)}/reports/${report.id}/updateDefinition`, "POST", { definition: { parts } });
}
