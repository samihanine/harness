/**
 * Screens are loaded when first opened (their own chunks). A load can fail for a reason that a new
 * page load fixes: the dev server restarted or re-optimized its dependencies while the tab stayed
 * open, a new version of the exported app replaced the chunks. So a failed load is tried again, then
 * the page is reloaded once; only a failure after that reaches the screen's error message.
 */
import { lazy, type ComponentType } from "react";

export const isLoadFailure = (error: unknown) => /dynamically imported module|Importing a module script failed|error loading dynamically imported|Loading chunk|Unable to preload/i.test(String((error as Error)?.message ?? error));

const RELOADED = "hub-reloaded-at";

async function load<T>(importer: () => Promise<T>): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    try {
      return await importer();
    } catch (error) {
      if (!isLoadFailure(error)) throw error;
      if (attempt < 2) {
        await new Promise((r) => setTimeout(r, 400 * (attempt + 1)));
        continue;
      }
      // Reloaded once already in the last minute: the failure is real, said as it is.
      let recent = false;
      try {
        recent = Date.now() - Number(sessionStorage.getItem(RELOADED) ?? 0) < 60_000;
        if (!recent) sessionStorage.setItem(RELOADED, String(Date.now()));
      } catch {}
      if (recent) throw error;
      location.reload();
      return new Promise<T>(() => undefined);
    }
  }
}

/** `lazy` with the retry above. */
export const lazyScreen = <T extends ComponentType<any>>(importer: () => Promise<{ default: T }>) => lazy(() => load(importer));
