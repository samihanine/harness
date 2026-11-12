/**
 * The slides of the deck as the app reads them (shapes with their places, ids, texts): for the
 * outlines on the pictures and the list of shapes. Never drawn from here: the pictures of the slides
 * come from Microsoft Office or the preview engine (rendered.ts).
 */
import { useEffect, useState } from "react";
import type { Deck, SlideInfo } from "./pptx/deck";

/** Every slide of the deck (read again when `version` changes). */
export function useSlides(deck: Deck | undefined, version: number) {
  const [slides, setSlides] = useState<SlideInfo[]>([]);
  useEffect(() => {
    let live = true;
    if (!deck) return setSlides([]);
    void deck.slides().then((list) => live && setSlides(list));
    return () => void (live = false);
  }, [deck, version]);
  return slides;
}
