/**
 * The deck's slides as pictures, drawn by the real thing — never rebuilt by the app:
 * 1. Microsoft Office: a work copy in the user's OneDrive converted to PDF (services/office-preview.ts),
 *    shown with pdf.js (pdf.ts). The PowerPoint rendering, for a local deck as for a SharePoint one.
 * 2. When that is not possible for this account (no OneDrive, no right, the option off in Settings):
 *    LibreOffice in the local server (services/office.ts), a close rendering.
 * Small pictures of every slide after each change, the slide shown at full size. The pictures of the
 * version before stay on screen until the new ones arrive. `problem` says why nothing could be drawn.
 */
import { useEffect, useRef, useState } from "react";
import { PRESENTER_OFFICE } from "@/config";
import { OfficeEngineError, renderDeck } from "@/services/office";
import { OfficePreviewError, officePdf } from "@/services/office-preview";
import type { Deck } from "./pptx/deck";
import type { Pdf } from "./pdf";

export type Rendered = {
  small: Record<number, string>;
  large: Record<number, string>;
  busy: boolean;
  /** Who drew the pictures. */
  source?: "office" | "libreoffice";
  problem?: { message: string; missing: boolean };
};

const SMALL = 360;
const LARGE = 1600;
/** Not tried again in this session once an account cannot use it. */
const off = { office: false, libreoffice: false };
const reason = (e: unknown) => (e instanceof Error ? e.message : String(e));

export function useRendered(deck: Deck | undefined, version: number, current: number, id = "deck"): Rendered {
  const [state, setState] = useState<Rendered>({ small: {}, large: {}, busy: false });
  /** What drew this version: the PDF (Office) or the bytes (LibreOffice), for the large pictures. */
  const drawn = useRef<{ key: string; pdf?: Pdf; bytes?: Uint8Array; fonts?: string[]; large: Map<number, string> } | undefined>(undefined);
  const key = deck ? `${id}:${version}` : "";

  useEffect(() => {
    if (!deck) {
      drawn.current = undefined;
      return setState({ small: {}, large: {}, busy: false });
    }
    let live = true;
    setState((s) => ({ ...s, busy: true }));
    // A short pause: several changes in a row are drawn once.
    const timer = setTimeout(() => {
      void (async () => {
        const bytes = await deck.bytes();
        const fonts = await deck.fonts();
        const problems: string[] = [];
        let missing = false;
        if (PRESENTER_OFFICE && !off.office) {
          try {
            const { openPdf, pageImage } = await import("./pdf");
            const pdf = await openPdf(await officePdf(id, bytes));
            const small: Record<number, string> = {};
            for (let n = 1; n <= pdf.numPages && live; n++) small[n] = await pageImage(pdf, n, SMALL);
            if (!live) return;
            drawn.current = { key, pdf, large: new Map() };
            return setState((s) => ({ ...s, small, large: {}, busy: false, source: "office", problem: undefined }));
          } catch (e) {
            if (e instanceof OfficePreviewError && e.unavailable) off.office = true;
            problems.push(reason(e));
          }
        }
        if (!off.libreoffice) {
          try {
            const r = await renderDeck(bytes, { width: SMALL, fonts });
            if (!live) return;
            drawn.current = { key, bytes, fonts, large: new Map() };
            return setState((s) => ({ ...s, small: Object.fromEntries(r.pages.map((p) => [p.page, p.url])), large: {}, busy: false, source: "libreoffice", problem: undefined }));
          } catch (e) {
            if (e instanceof OfficeEngineError && e.missing) {
              off.libreoffice = true;
              missing = true;
            }
            problems.push(reason(e));
          }
        }
        if (live) setState({ small: {}, large: {}, busy: false, problem: { message: problems.join(" — "), missing } });
      })();
    }, 500);
    return () => {
      live = false;
      clearTimeout(timer);
    };
  }, [deck, key]); // eslint-disable-line react-hooks/exhaustive-deps

  // The slide shown, at full size.
  useEffect(() => {
    let live = true;
    const mine = drawn.current;
    if (!deck || !mine || mine.key !== key) return;
    const have = mine.large.get(current);
    if (have) return void setState((s) => ({ ...s, large: { ...s.large, [current]: have } }));
    void (async () => {
      let url: string | undefined;
      if (mine.pdf) {
        const { pageImage } = await import("./pdf");
        url = await pageImage(mine.pdf, Math.min(current, mine.pdf.numPages), LARGE);
      } else if (mine.bytes) url = (await renderDeck(mine.bytes, { width: LARGE, pages: [current], fonts: mine.fonts })).pages[0]?.url;
      if (!url || drawn.current !== mine) return;
      mine.large.set(current, url);
      if (live) setState((s) => ({ ...s, large: { ...s.large, [current]: url! } }));
    })().catch(() => undefined);
    return () => void (live = false);
  }, [deck, key, current, state.source]); // eslint-disable-line react-hooks/exhaustive-deps

  return state;
}
