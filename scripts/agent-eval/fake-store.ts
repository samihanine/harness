/** In-memory stand-in for src/lib/store.ts (same API), seeded by each scenario. */
/* Loaded in place of src/lib/store.ts (see preload.ts): its types are the real file's, for the compiler only. */
type Settings = { aiKey?: string; model?: string };

const collections = new Map<string, { id: string }[]>();

export const reset = (seed: Record<string, any[]> = {}) => {
  collections.clear();
  for (const [k, v] of Object.entries(seed)) collections.set(k, structuredClone(v));
};
/** Synchronous read (for agents whose getters are synchronous, like the screens' React state). */
export const peek = (name: string): any[] => collections.get(name) ?? [];
export const all = async (name: string): Promise<any[]> => collections.get(name) ?? [];
export const put = async (name: string, item: any) => {
  const items = await all(name);
  collections.set(name, items.some((i) => i.id === item.id) ? items.map((i) => (i.id === item.id ? item : i)) : [...items, item]);
  return item;
};
export const patch = async (name: string, id: string, values: any) => void collections.set(name, (await all(name)).map((i) => (i.id === id ? { ...i, ...values } : i)));
export const remove = async (name: string, id: string) => void collections.set(name, (await all(name)).filter((i) => i.id !== id));
export const find = async (name: string, id?: string) => (id ? (await all(name)).find((i) => i.id === id) : undefined);
export const useCollection = () => [];
export const persist = () => undefined;
export const settings = {
  get: async (): Promise<Settings> => ({ aiKey: process.env.AI_KEY, model: process.env.AI_MODEL }),
  set: async () => undefined,
};
