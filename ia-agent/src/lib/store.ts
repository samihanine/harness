/** Persistence in IndexedDB (conversations with their traces can be large). */
import { createStore, del, get, set } from "idb-keyval";
import type { Conversation } from "./types";

const db = createStore("agent", "kv");

export type Settings = {
  aiKey: string;
  model: string;
  /** "auto": changes are applied directly (and can be undone). "ask": each change waits for approval. */
  approval: "auto" | "ask";
};

const defaults: Settings = { aiKey: "", model: "", approval: "auto" };

export const settings = {
  get: async (): Promise<Settings> => ({ ...defaults, ...(await get("settings", db)) }),
  set: async (patch: Partial<Settings>) => set("settings", { ...(await settings.get()), ...patch }, db),
};

type Meta = Omit<Conversation, "items">;

export const conversations = {
  list: async (scope: string) =>
    ((await get<Meta[]>("conversations", db)) ?? [])
      .filter((c) => c.scope === scope)
      .sort((a, b) => b.updatedAt - a.updatedAt),
  get: (id: string) => get<Conversation>(`conversation:${id}`, db),
  async put(conversation: Conversation) {
    await set(`conversation:${conversation.id}`, conversation, db);
    const { items: _, ...meta } = conversation;
    const all = ((await get<Meta[]>("conversations", db)) ?? []).filter((c) => c.id !== conversation.id);
    await set("conversations", [...all, meta], db);
  },
  async remove(id: string) {
    await del(`conversation:${id}`, db);
    const all = (await get<Meta[]>("conversations", db)) ?? [];
    await set("conversations", all.filter((c) => c.id !== id), db);
  },
};
