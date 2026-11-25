/** Semantic models as the AI knows them: read once (TMDL or DAX INFO views, with value samples), kept 15 minutes. */
import type { Ref } from "@shared/types";
import { readModel, type ModelInfo } from "./model";

const cache = new Map<string, { at: number; info: Promise<ModelInfo> }>();

export function modelInfo(ref: Ref, fresh = false): Promise<ModelInfo> {
  const hit = cache.get(ref.id);
  if (hit && !fresh && Date.now() - hit.at < 15 * 60_000) return hit.info;
  const info = readModel(ref, true);
  cache.set(ref.id, { at: Date.now(), info });
  info.catch(() => cache.delete(ref.id));
  return info;
}

export const forgetModel = (id: string) => cache.delete(id);

/** A model the browser already read (kept with its item): used as is while it is fresh enough. */
export function seedModel(id: string, info: unknown) {
  if (!info || typeof info !== "object" || !Array.isArray((info as ModelInfo).tables)) return;
  const hit = cache.get(id);
  if (hit && Date.now() - hit.at < 15 * 60_000) return;
  cache.set(id, { at: Date.now(), info: Promise.resolve(info as ModelInfo) });
}
