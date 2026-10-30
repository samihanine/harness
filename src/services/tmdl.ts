/** Semantic models as TMDL through the Fabric API (read and write their definition). */
import { fabric } from "./ms";

export type ModelFile = { path: string; text: string };

/** Fabric needs a workspace id: "My workspace" has one too. */
export async function workspaceId(groupId?: string) {
  if (groupId) return groupId;
  const { value } = await fabric("/workspaces");
  return value.find((w: { type: string }) => w.type === "Personal").id as string;
}

const encode = (files: ModelFile[]) => files.map((f) => ({ path: f.path, payload: btoa(String.fromCharCode(...new TextEncoder().encode(f.text))), payloadType: "InlineBase64" }));
const decode = (payload: string) => new TextDecoder().decode(Uint8Array.from(atob(payload), (c) => c.charCodeAt(0)));

export async function readModelFiles(datasetId: string, groupId?: string): Promise<ModelFile[]> {
  const result = await fabric(`/workspaces/${await workspaceId(groupId)}/semanticModels/${datasetId}/getDefinition?format=TMDL`, "POST");
  return result.definition.parts.filter((p: { path: string }) => p.path !== ".platform").map((p: { path: string; payload: string }) => ({ path: p.path, text: decode(p.payload) }));
}

export async function writeModelFiles(datasetId: string, files: ModelFile[], groupId?: string) {
  await fabric(`/workspaces/${await workspaceId(groupId)}/semanticModels/${datasetId}/updateDefinition`, "POST", { definition: { parts: encode(files) } });
}
