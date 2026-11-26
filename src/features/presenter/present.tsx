/**
 * The Present tab: every slide in a gallery, and the presentation in full screen (arrows, Esc). The Power BI
 * views of the slides are live in full screen — the real report, interactive — over their place on the slide;
 * the PowerPoint keeps pictures of them.
 */
import { useEffect, useRef, useState } from "react";
import { ChevronLeftIcon, ChevronRightIcon, PlayIcon, XIcon } from "lucide-react";
import { resolveSlide, SLIDE, templateOf, type SlideContent, type Template } from "@shared/slides";
import type { Ref } from "@shared/types";
import { Button } from "@/components/ui/button";
import { embedToken, embedUrl, powerbi } from "@/features/powerbi/embed";
import { post } from "@/lib/api";
import { SlideView } from "./slides";

type Resolved = { url: string; ref?: Ref; pageName?: string; visualName?: string };

/** One live Power BI view: a visual alone, or a report page without its panes. */
function LiveZone({ view, box, scale }: { view: Resolved; box: { x: number; y: number; w: number; h: number }; scale: number }) {
  const el = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const div = el.current;
    if (!div || !view.ref) return;
    let gone = false;
    void (async () => {
      const { pbi, service } = await powerbi();
      if (gone) return;
      const common = { id: view.ref!.id, embedUrl: embedUrl(view.ref!), accessToken: await embedToken(), tokenType: pbi.models.TokenType.Aad, settings: { panes: { filters: { visible: false }, pageNavigation: { visible: false } }, background: pbi.models.BackgroundType.Transparent } };
      service.embed(div, view.visualName && view.pageName ? ({ ...common, type: "visual", pageName: view.pageName, visualName: view.visualName } as never) : ({ ...common, type: "report", pageName: view.pageName, settings: { ...common.settings, layoutType: pbi.models.LayoutType.Custom, customLayout: { displayOption: pbi.models.DisplayOption.FitToPage } } } as never));
    })();
    return () => {
      gone = true;
      void powerbi().then(({ service }) => service.reset(div));
    };
  }, [view.ref?.id, view.pageName, view.visualName]); // eslint-disable-line react-hooks/exhaustive-deps
  const px = SLIDE.dpi * scale;
  return <div ref={el} className="absolute overflow-hidden bg-white [&_iframe]:border-0" style={{ left: box.x * px, top: box.y * px, width: box.w * px, height: box.h * px }} />;
}

function FullScreen({ slides, templates, start, onClose }: { slides: SlideContent[]; templates: Template[]; start: number; onClose: () => void }) {
  const [i, setI] = useState(start);
  const [size, setSize] = useState({ w: 0, h: 0 });
  const [views, setViews] = useState<Record<string, Resolved>>({});
  const root = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = root.current!;
    void el.requestFullscreen?.().catch(() => undefined);
    const key = (e: KeyboardEvent) => {
      if (e.key === "ArrowRight" || e.key === " " || e.key === "PageDown") setI((x) => Math.min(slides.length - 1, x + 1));
      if (e.key === "ArrowLeft" || e.key === "PageUp") setI((x) => Math.max(0, x - 1));
      if (e.key === "Escape") onClose();
    };
    const resize = () => setSize({ w: window.innerWidth, h: window.innerHeight });
    const exit = () => !document.fullscreenElement && onClose();
    resize();
    window.addEventListener("keydown", key);
    window.addEventListener("resize", resize);
    document.addEventListener("fullscreenchange", exit);
    // The Power BI links of the deck: their report, page and visual names, for the live views.
    const urls = [...new Set(slides.flatMap((c) => c.powerbi))];
    if (urls.length) void post<Resolved[]>("/slides/resolve", { urls }).then((r) => setViews(Object.fromEntries(r.map((x) => [x.url, x]))), () => undefined);
    return () => {
      window.removeEventListener("keydown", key);
      window.removeEventListener("resize", resize);
      document.removeEventListener("fullscreenchange", exit);
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  const w = Math.min(size.w, (size.h * 16) / 9);
  const scale = w / 1280;
  const c = slides[i];
  const t = templateOf(templates, c.template);
  const zones = resolveSlide(t, c).filter((d) => d.type === "powerbi") as { url: string; x: number; y: number; w: number; h: number }[];
  return (
    <div ref={root} className="fixed inset-0 z-50 flex items-center justify-center bg-black">
      <div className="relative" style={{ width: w, height: (w * 9) / 16 }}>
        <SlideView template={t} content={c} className="absolute inset-0" live />
        {zones.map((z, k) => views[z.url]?.ref && <LiveZone key={`${i}-${k}`} view={views[z.url]} box={z} scale={scale} />)}
      </div>
      <div className="absolute right-4 bottom-4 flex items-center gap-1 rounded-full bg-white/10 px-2 py-1 text-white opacity-30 transition-opacity hover:opacity-100">
        <button onClick={() => setI(Math.max(0, i - 1))} className="p-1.5" aria-label="Previous"><ChevronLeftIcon className="size-5" /></button>
        <span className="text-[13px] tabular-nums">{i + 1} / {slides.length}</span>
        <button onClick={() => setI(Math.min(slides.length - 1, i + 1))} className="p-1.5" aria-label="Next"><ChevronRightIcon className="size-5" /></button>
        <button onClick={onClose} className="p-1.5" aria-label="Exit"><XIcon className="size-5" /></button>
      </div>
    </div>
  );
}

export function Present({ slides, templates }: { slides: SlideContent[]; templates: Template[] }) {
  const [playing, setPlaying] = useState<number>();
  return (
    <div className="flex flex-col gap-4 p-6">
      <div className="flex items-center gap-2">
        <div className="flex-1 text-[13px] text-muted-foreground">{slides.length} slides{slides.some((c) => c.powerbi.length) ? " · Power BI views live in full screen" : ""}</div>
        <Button variant="dark" size="sm" disabled={!slides.length} onClick={() => setPlaying(0)}><PlayIcon /> Present</Button>
      </div>
      <div className="grid gap-4 sm:grid-cols-2 2xl:grid-cols-3">
        {slides.map((c, i) => (
          <button key={i} onClick={() => setPlaying(i)} className="group text-left" title="Present from this slide">
            <SlideView template={templateOf(templates, c.template)} content={c} className="rounded-xl border shadow-[0_1px_4px_rgba(0,0,0,0.06)] transition-transform group-hover:-translate-y-0.5" />
            <div className="mt-1 text-[12px] text-muted-foreground">{i + 1}</div>
          </button>
        ))}
      </div>
      {playing !== undefined && <FullScreen slides={slides} templates={templates} start={playing} onClose={() => setPlaying(undefined)} />}
    </div>
  );
}
