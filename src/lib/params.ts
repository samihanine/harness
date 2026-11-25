/**
 * What a screen shows, in its address: the item (?item=…) and the conversation (?c=…). Each browser tab has
 * its own, so several items and conversations of the same assistant can be open (and running) side by side;
 * without them, the last item chosen in this browser.
 */
import { useCallback } from "react";
import { useNavigate, useSearch } from "@tanstack/react-router";
import { useLocal } from "./store";

export type ScreenSearch = { item?: string; c?: string };
const text = (v: unknown) => (typeof v === "string" && v ? v : undefined);
export const screenSearch = (s: Record<string, unknown>): ScreenSearch => ({ item: text(s.item), c: text(s.c) });

export function useScreenSearch() {
  const search = useSearch({ strict: false }) as ScreenSearch;
  const navigate = useNavigate();
  const set = useCallback((patch: ScreenSearch) => void navigate({ to: ".", search: ((prev: ScreenSearch) => ({ ...prev, ...patch })) as never, replace: true }), [navigate]);
  return [search, set] as const;
}

/** The item of a screen: the address's, else the last one chosen here; choosing writes both. */
export function useItemParam(key: string): [string, (id: string) => void] {
  const [search, set] = useScreenSearch();
  const [last, setLast] = useLocal(key, "");
  const choose = useCallback((next: string) => (setLast(next), set({ item: next || undefined, c: undefined })), [set]); // eslint-disable-line react-hooks/exhaustive-deps
  return [search.item ?? last, choose];
}
