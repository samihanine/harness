/**
 * The Builder's draft: the report's definition (PBIR) read once per conversation, changed by the tools,
 * written back whole by save_report (then the embedded report is reloaded). Unsaved changes stay in the
 * conversation; discard_changes reads the report again.
 */
import type { Ref } from "@shared/types";
import { anyReportDefinition, partText, reportDefinition, textPart, writeReportDefinition, type Part } from "./definitions";
import { isLegacy, legacyToPbir, pbirToLegacy, type Legacy } from "./legacy";
import { modelInfo } from "./models";
import type { ModelInfo } from "./model";

/** `legacy`: the report is saved in the old format (shown as PBIR files, written back into its report.json). */
export type Draft = { ref: Ref; parts: Part[]; changed: Set<string>; readAt: number; legacy?: Legacy };

/** The definition as PBIR files, whatever format the report is saved in. */
export async function definitionFiles(ref: Ref): Promise<{ parts: Part[]; legacy?: Legacy }> {
  try {
    return { parts: await reportDefinition(ref) };
  } catch (e) {
    const any = await anyReportDefinition(ref).catch(() => {
      throw e;
    });
    if (!isLegacy(any)) throw e;
    return legacyToPbir(any);
  }
}

export async function draftOf(scratch: Record<string, any>, ref: Ref): Promise<Draft> {
  const d = scratch.draft as Draft | undefined;
  if (d && d.ref.id === ref.id) return d;
  const draft: Draft = { ref, ...(await definitionFiles(ref)), changed: new Set(), readAt: Date.now() };
  scratch.draft = draft;
  return draft;
}

export const readJson = (d: Draft, path: string) => {
  const p = d.parts.find((x) => x.path === path);
  if (!p) throw new Error(`No file ${path} in the report.`);
  return JSON.parse(partText(p));
};

export function writeJson(d: Draft, path: string, value: unknown) {
  const text = JSON.stringify(value, null, 2);
  const i = d.parts.findIndex((x) => x.path === path);
  const part = textPart(path, text);
  if (i >= 0) d.parts[i] = part;
  else d.parts.push(part);
  d.changed.add(path);
}

export function removePath(d: Draft, prefix: string) {
  const before = d.parts.length;
  d.parts = d.parts.filter((p) => !(p.path === prefix || p.path.startsWith(`${prefix}/`)));
  if (d.parts.length !== before) d.changed.add(prefix);
}

/** Pages in order: name, display name, size. */
export function pagesOf(d: Draft) {
  const order: string[] = d.parts.find((p) => p.path === "definition/pages/pages.json") ? readJson(d, "definition/pages/pages.json").pageOrder ?? [] : [];
  return d.parts
    .filter((p) => /^definition\/pages\/[^/]+\/page\.json$/.test(p.path))
    .map((p) => JSON.parse(partText(p)))
    .sort((a, b) => (order.indexOf(a.name) + 1 || 999) - (order.indexOf(b.name) + 1 || 999)) as { name: string; displayName: string; width: number; height: number }[];
}

export function pageOf(d: Draft, page: string) {
  const pages = pagesOf(d);
  const p = pages.find((x) => x.name === page) ?? pages.find((x) => x.displayName.toLowerCase() === page.toLowerCase());
  if (!p) throw new Error(`No page "${page}". Pages: ${pages.map((x) => `"${x.displayName}" (${x.name})`).join(", ")}`);
  return p;
}

export function visualsOf(d: Draft, page: string) {
  return d.parts.filter((p) => p.path.startsWith(`definition/pages/${page}/visuals/`) && p.path.endsWith("/visual.json")).map((p) => ({ path: p.path, json: JSON.parse(partText(p)) }));
}

const titleOf = (v: any) => String(v.visual?.visualContainerObjects?.title?.[0]?.properties?.text?.expr?.Literal?.Value ?? "").replace(/^'(.*)'$/s, "$1");

export function visualOf(d: Draft, page: string, visual: string) {
  const list = visualsOf(d, page);
  const hit = list.find((v) => v.json.name === visual) ?? list.filter((v) => titleOf(v.json).toLowerCase() === visual.toLowerCase())[0];
  if (!hit) throw new Error(`No visual "${visual}" on page ${page}. Visuals: ${list.map((v) => `${v.json.visual?.visualType} "${titleOf(v.json)}" (${v.json.name})`).join(", ") || "none"}`);
  return hit;
}

/** Report measures (reportExtensions.json) for field references. */
export function reportMeasures(d: Draft) {
  const p = d.parts.find((x) => x.path === "definition/reportExtensions.json");
  if (!p) return [] as { table: string; name: string }[];
  return (JSON.parse(partText(p)).entities ?? []).flatMap((e: any) => (e.measures ?? []).map((m: any) => ({ table: e.name, name: m.name })));
}

export async function draftModel(ctxReport: { datasetId?: string; datasetGroupId?: string }): Promise<ModelInfo | undefined> {
  if (!ctxReport.datasetId) return undefined;
  return modelInfo({ id: ctxReport.datasetId, groupId: ctxReport.datasetGroupId }).catch(() => undefined);
}

/** The $schema the report's visuals use (new visuals take the same version). */
export const visualSchema = (d: Draft) => {
  const any = d.parts.find((p) => p.path.endsWith("/visual.json"));
  return any ? JSON.parse(partText(any)).$schema : undefined;
};

export async function saveDraft(d: Draft) {
  await writeReportDefinition(d.ref, d.legacy ? pbirToLegacy(d.parts, d.legacy, d.changed) : d.parts);
  // The next changes compare with what is saved now.
  if (d.legacy) Object.assign(d, await definitionFiles(d.ref));
  const n = d.changed.size;
  d.changed.clear();
  return n;
}

/** Applies "add" JSON Patch operations (what formatOps returns). */
export function applyAdds(target: any, ops: { op: "add"; path: string; value: unknown }[]) {
  for (const op of ops) {
    const keys = op.path.split("/").slice(1).map((k) => k.replace(/~1/g, "/").replace(/~0/g, "~"));
    let cur = target;
    keys.slice(0, -1).forEach((k, i) => {
      const nextIsIndex = /^\d+$/.test(keys[i + 1]);
      if (cur[k] === undefined) cur[k] = nextIsIndex ? [] : {};
      cur = cur[k];
    });
    const last = keys.at(-1)!;
    if (Array.isArray(cur) && /^\d+$/.test(last) && Number(last) >= cur.length) cur.push(op.value);
    else cur[last] = op.value;
  }
  return target;
}
