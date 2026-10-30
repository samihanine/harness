/** Report definition (PBIR) through Fabric: read and written whole (the editor's draft works on it, see report-draft.ts). */
import { fabric } from "./ms";
import type { Ref } from "./pbi";
import { workspaceId } from "./tmdl";

export type Part = { path: string; payload: string; payloadType: string };
const decode = (p: string) => new TextDecoder().decode(Uint8Array.from(atob(p), (c) => c.charCodeAt(0)));
const encode = (s: string) => btoa(String.fromCharCode(...new TextEncoder().encode(s)));

export async function definition(report: Ref): Promise<Part[]> {
  const ws = await workspaceId(report.groupId);
  const result = await fabric(`/workspaces/${ws}/reports/${report.id}/getDefinition?format=PBIR`, "POST").catch((e) => {
    throw new Error(`The report definition is not readable in PBIR format (${e instanceof Error ? e.message : e}). Save the report once from the editor, and check you can edit it.`);
  });
  return result.definition.parts.filter((p: Part) => p.path !== ".platform");
}

/**
 * The definition in the report's own format when PBIR is refused (reports still saved in the older
 * "PBIR-Legacy" format: one report.json, which the API does not convert).
 */
export async function legacyDefinition(report: Ref): Promise<Part[]> {
  const ws = await workspaceId(report.groupId);
  const result = await fabric(`/workspaces/${ws}/reports/${report.id}/getDefinition`, "POST");
  return result.definition.parts.filter((p: Part) => p.path !== ".platform");
}

/** A part's text, and a text as a part's payload. */
export const partText = (p: Part) => decode(p.payload);
export const textPayload = (text: string) => encode(text);

/** Writes a whole definition (all its parts) back to the report. */
export async function writeDefinition(report: Ref, parts: Part[]) {
  await fabric(`/workspaces/${await workspaceId(report.groupId)}/reports/${report.id}/updateDefinition`, "POST", { definition: { parts } });
}
