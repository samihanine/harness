/**
 * A slide drawn in the app from its XML (deck.ts): what its layout and master draw, then its shapes —
 * texts at their size, pictures, fills — at their places. A preview to follow the changes (the AI's
 * and the user's), not a PowerPoint renderer: effects, theme colors and charts show as frames.
 */
import { useEffect, useState } from "react";
import { cn } from "@/lib/utils";
import type { Deck, Shape, SlideInfo } from "./pptx/deck";

export type DrawnSlide = SlideInfo & { chrome: Shape[]; base?: string };

/** Every slide of the deck with what it shows (read again when `version` changes). */
export function useSlides(deck: Deck | undefined, version: number) {
  const [slides, setSlides] = useState<DrawnSlide[]>([]);
  useEffect(() => {
    let live = true;
    if (!deck) return setSlides([]);
    void (async () => {
      const out: DrawnSlide[] = [];
      for (const s of await deck.slides()) {
        const chrome = await deck.chromeOf(s.number);
        out.push({ ...s, chrome: chrome.shapes, base: chrome.background });
      }
      if (live) setSlides(out);
    })();
    return () => void (live = false);
  }, [deck, version]);
  return slides;
}

/** Pictures of the deck as URLs (read once per part). */
const urls = new WeakMap<Deck, Map<string, Promise<string>>>();
function usePicture(deck: Deck, part?: string) {
  const [url, setUrl] = useState<string>();
  useEffect(() => {
    if (!part) return;
    if (!urls.has(deck)) urls.set(deck, new Map());
    const cache = urls.get(deck)!;
    if (!cache.has(part)) cache.set(part, deck.imageUrl(part).catch(() => ""));
    void cache.get(part)!.then(setUrl);
  }, [deck, part]);
  return url;
}

function ShapeView({ deck, shape, highlight }: { deck: Deck; shape: Shape; highlight?: boolean }) {
  const url = usePicture(deck, shape.image);
  if (!shape.box) return null;
  const { w, h } = deck.size;
  const widthPt = w / 12_700;
  const style = { left: `${(shape.box.x / w) * 100}%`, top: `${(shape.box.y / h) * 100}%`, width: `${(shape.box.w / w) * 100}%`, height: `${(shape.box.h / h) * 100}%`, background: shape.fill };
  const pt = (size: number) => `${(size / widthPt) * 100}cqw`;
  return (
    <div style={style} title={`#${shape.id} ${shape.name}`} className={cn("absolute overflow-hidden", highlight && "outline-2 outline-primary", shape.kind === "chart" && "flex items-center justify-center border border-dashed text-[1.4cqw] text-muted-foreground", shape.kind === "line" && "border-t border-zinc-400")}>
      {shape.kind === "picture" && url && <img src={url} alt="" className="h-full w-full object-fill" />}
      {shape.kind === "chart" && "chart"}
      {shape.kind === "table" && <div className="grid h-full grid-cols-3 border border-zinc-300 text-[1cqw]">{shape.paragraphs.slice(0, 24).map((p, i) => <span key={i} className="truncate border border-zinc-200 px-[0.3cqw]">{p.text}</span>)}</div>}
      {(shape.kind === "text" || shape.kind === "other") && (
        <div className={cn("flex h-full flex-col leading-[1.15]", shape.placeholder === "title" || shape.placeholder === "ctrTitle" ? "justify-center" : "justify-start")}>
          {shape.paragraphs.map((p, i) => (
            <p key={i} style={{ fontSize: pt(p.size ?? shape.size), paddingLeft: `${p.level * 2}cqw`, color: p.color, fontWeight: p.bold ? 700 : undefined, textAlign: p.align === "ctr" ? "center" : p.align === "r" ? "right" : undefined }} className="min-h-[0.6em] break-words whitespace-pre-wrap">
              {p.level > 0 || (shape.placeholder === "body" && p.text) ? "• " : ""}
              {p.text}
            </p>
          ))}
        </div>
      )}
    </div>
  );
}

export function SlideView({ deck, slide, className, highlight }: { deck: Deck; slide: DrawnSlide; className?: string; highlight?: string }) {
  return (
    <div className={cn("@container relative w-full overflow-hidden bg-white text-zinc-900 shadow-sm ring-1 ring-border", className)} style={{ aspectRatio: `${deck.size.w} / ${deck.size.h}`, background: slide.background ?? slide.base ?? "#ffffff" }}>
      {slide.chrome.map((s, i) => (
        <ShapeView key={`c${i}`} deck={deck} shape={s} />
      ))}
      {slide.shapes.map((s) => (
        <ShapeView key={s.id} deck={deck} shape={s} highlight={highlight === s.id} />
      ))}
    </div>
  );
}
