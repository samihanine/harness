/** Pictures of a deck's slides as PowerPoint draws them (the server renders them through Microsoft), kept per version. */
import { useEffect, useRef, useState } from "react";
import { errorText, post } from "@/lib/api";

const cache = new Map<string, string[]>();

export type SlidesSource = { markdown: string } | { fileId: string };

export function useSlides(source: SlidesSource | undefined, key: string) {
  const [pics, setPics] = useState<string[] | undefined>(cache.get(key));
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string>();
  const asked = useRef("");
  const render = async () => {
    if (!source) return;
    setBusy(true);
    setError(undefined);
    asked.current = key;
    try {
      const ids = await post<string[]>("/deck/preview", source);
      cache.set(key, ids);
      if (asked.current === key) setPics(ids);
    } catch (e) {
      if (asked.current === key) setError(errorText(e));
    } finally {
      if (asked.current === key) setBusy(false);
    }
  };
  useEffect(() => {
    const hit = cache.get(key);
    if (hit) return setPics(hit);
    // The pictures of the previous version stay while the new ones are drawn.
    void render();
  }, [key]); // eslint-disable-line react-hooks/exhaustive-deps
  return { pics, busy, error, render };
}
