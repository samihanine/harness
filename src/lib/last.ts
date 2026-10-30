/** Last selection of each screen (report, Excel file, conversation), remembered in this browser. */
import { useEffect } from "react";

export const last = {
  get: (key: string) => {
    try {
      return localStorage.getItem(`last:${key}`) ?? undefined;
    } catch {
      return undefined;
    }
  },
  set: (key: string, value?: string) => {
    try {
      if (value) localStorage.setItem(`last:${key}`, value);
    } catch {}
  },
};

/**
 * Restores the last value of a search param when the screen is opened without a selection
 * (`opened` = something else is selected), and remembers it when it changes.
 */
export function useLast(key: string, value: string | undefined, restore: (value: string) => void, opened = !!value) {
  useEffect(() => {
    const saved = last.get(key);
    if (!opened && saved) restore(saved);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => last.set(key, value), [key, value]);
}
