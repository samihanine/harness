/**
 * Reports and models as the screens use them: their structure read once when added (and on "Refresh"),
 * kept in this browser and sent with the conversation, so the assistant never waits to read it again.
 */
import type { DatasetRef, ReportRef } from "@shared/types";
import { get } from "./api";
import { item, list, save } from "./store";

const pending = new Map<string, Promise<unknown>>();
const once = <T,>(key: string, f: () => Promise<T>) => {
  if (!pending.has(key)) pending.set(key, f().finally(() => pending.delete(key)));
  return pending.get(key) as Promise<T>;
};

const query = (r: { id: string; groupId?: string; appId?: string }, extra: Record<string, string> = {}) => new URLSearchParams({ id: r.id, ...(r.groupId ? { groupId: r.groupId } : {}), ...(r.appId ? { appId: r.appId } : {}), ...extra });

/** A report's structure (read now when missing or when `fresh`), saved with the report. */
export function reportStructure(r: ReportRef, fresh = false) {
  return once(`structure:${r.id}`, async () => {
    const kept = await item("report", r.id);
    if (kept?.structure && !fresh) return kept.structure;
    const structure = await get(`/powerbi/report-structure?${query(r)}`);
    const now = await item("report", r.id);
    if (now) await save("report", { ...now, structure, structureAt: new Date().toISOString() });
    return structure;
  });
}

/** A model's structure (with value examples), read once and kept. */
export function modelInfo(ref: { id: string; groupId?: string }, fresh = false) {
  return once(`model:${ref.id}`, async () => {
    const kept = await item("model", ref.id);
    if (kept && !fresh) return kept.info;
    const info = await get(`/powerbi/model?${query(ref, fresh ? { fresh: "1" } : {})}`);
    await save("model", { id: ref.id, info, at: new Date().toISOString() });
    return info;
  });
}

/** Reads what a new report or model needs, in the background. */
export function prepareReport(r: ReportRef) {
  void reportStructure(r).catch(() => undefined);
  if (r.datasetId) void modelInfo({ id: r.datasetId, groupId: r.datasetGroupId }).catch(() => undefined);
}
export const prepareDataset = (d: DatasetRef) => void modelInfo(d).catch(() => undefined);

/** The models kept for these ids (sent with the conversation). */
export async function modelsFor(ids: (string | undefined)[]) {
  const all = await list("model");
  return Object.fromEntries(all.filter((m) => ids.includes(m.id)).map((m) => [m.id, m.info]));
}

/** A report with its kept structure (the store's copy). */
export const withStructure = async (r?: ReportRef) => (r ? ((await item("report", r.id)) ?? r) : undefined);
