/**
 * The deck's slides as LibreOffice draws them (services/office.ts): small pictures of every slide
 * after each change, and the slide shown at full size. The pictures of the version before stay on
 * screen until the new ones arrive. When the engine is missing or fails, `problem` says why and the
 * screen falls back to its own drawing (preview.tsx).
 */
import { useEffect, useRef, useState } from "react";
import { OfficeEngineError, renderDeck } from "@/services/office";
import type { Deck } from "./pptx/deck";

export type Rendered = { small: Record<number, string>; large: Record<number, string>; busy: boolean; problem?: { message: string; missing: boolean } };

const SMALL = 360;
const LARGE = 1600;

export function useRendered(deck: Deck | undefined, version: number, current: number): Rendered {
  const [state, setState] = useState<Rendered>({ small: {}, large: {}, busy: false });
  /** The deck's bytes and fonts of this version, made once for both sizes. */
  const input = useRef<{ key: string; data: Promise<{ bytes: Uint8Array; fonts: string[] }> } | undefined>(undefined);
  const drawn = useRef(new Map<number, string>());
  const off = useRef(false);
  const key = deck ? `${version}` : "";

  useEffect(() => {
    let live = true;
    drawn.current = new Map();
    if (!deck || off.current) return void setState((s) => ({ ...s, small: deck ? s.small : {}, large: deck ? s.large : {}, busy: false }));
    input.current = { key, data: (async () => ({ bytes: await deck.bytes(), fonts: await deck.fonts() }))() };
    setState((s) => ({ ...s, busy: true }));
    void input.current.data
      .then(({ bytes, fonts }) => renderDeck(bytes, { width: SMALL, fonts }))
      .then(
        (r) => live && setState((s) => ({ ...s, small: Object.fromEntries(r.pages.map((p) => [p.page, p.url])), busy: false, problem: undefined })),
        (e) => {
          // Not installed: not asked again in this session.
          if (e instanceof OfficeEngineError && e.missing) off.current = true;
          if (live) setState({ small: {}, large: {}, busy: false, problem: { message: e instanceof Error ? e.message : String(e), missing: e instanceof OfficeEngineError && e.missing } });
        },
      );
    return () => void (live = false);
  }, [deck, key]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    let live = true;
    const mine = input.current;
    if (!deck || off.current || !mine || mine.key !== key) return;
    if (drawn.current.has(current)) return void setState((s) => ({ ...s, large: { ...s.large, [current]: drawn.current.get(current)! } }));
    void mine.data
      .then(({ bytes, fonts }) => renderDeck(bytes, { width: LARGE, pages: [current], fonts }))
      .then((r) => {
        const url = r.pages[0]?.url;
        if (!url || input.current !== mine) return;
        drawn.current.set(current, url);
        if (live) setState((s) => ({ ...s, large: { ...s.large, [current]: url } }));
      }, () => undefined);
    return () => void (live = false);
  }, [deck, key, current]); // eslint-disable-line react-hooks/exhaustive-deps

  return state;
}
