/**
 * The Presenter's drawing in the app: a slide from its template and content (the same HTML the server pictures
 * and prints), scaled to its box; the browser's templates (the company's to start with, then the user's).
 */
import { useEffect, useMemo, useRef, useState } from "react";
import { slideHtml, type SlideContent, type Template } from "@shared/slides";
import { DEFAULT_TEMPLATES } from "@shared/slide-templates";
import { list, save, useItems } from "@/lib/store";

/** Where a picture of a slide is read from in the app. */
export const srcUrl = (src: string) => (src.startsWith("asset:") ? `/api/assets/template/${src.slice(6)}` : src.startsWith("file:") ? `/api/files/${src.slice(5)}` : src);
export const powerbiUrl = (url: string) => `/api/slides/powerbi?url=${encodeURIComponent(url)}`;

/** The templates of this browser, the company's added the first time. */
let seeded: Promise<void> | undefined;
export function useTemplates() {
  const items = useItems("template");
  useEffect(() => {
    seeded ??= list("template").then(async (all) => {
      if (!all.length) for (const t of DEFAULT_TEMPLATES) await save("template", t);
    });
  }, []);
  // The company's templates until this browser's are ready.
  return useMemo(() => [...(items?.length ? items : DEFAULT_TEMPLATES)].sort((a, b) => Number(!!b.favorite) - Number(!!a.favorite) || a.name.localeCompare(b.name)), [items]);
}

/** One slide drawn at the width of its box (16:9). */
export function SlideView({ template, content, className, live }: { template: Template; content: SlideContent; className?: string; live?: boolean }) {
  const box = useRef<HTMLDivElement>(null);
  const [width, setWidth] = useState(0);
  useEffect(() => {
    const el = box.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setWidth(el.clientWidth));
    ro.observe(el);
    setWidth(el.clientWidth);
    return () => ro.disconnect();
  }, []);
  const html = useMemo(() => slideHtml(template, content, { src: srcUrl, powerbi: (u) => (live ? undefined : powerbiUrl(u)), scale: width / 1280 }), [template, content, width, live]);
  return <div ref={box} className={className} style={{ aspectRatio: "16 / 9", overflow: "hidden", position: "relative" }} dangerouslySetInnerHTML={{ __html: width ? html : "" }} />;
}
