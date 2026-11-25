/**
 * The Present tab: every slide in a gallery, and the presentation in full screen (arrows, Esc). In a deck
 * made from the template, the Power BI views of its slides are live in full screen — the real report,
 * interactive — over their place on the slide; the PowerPoint keeps pictures of them.
 */
import { useEffect, useRef, useState } from "react";
import { ChevronLeftIcon, ChevronRightIcon, Loader2Icon, PlayIcon, RefreshCwIcon, XIcon } from "lucide-react";
import { SLIDE, type PbiZone } from "@shared/deck";
import type { Ref } from "@shared/types";
import { ErrorBox } from "@/components/common";
import { Button } from "@/components/ui/button";
import { embedToken, embedUrl, powerbi } from "@/features/powerbi/embed";
import { fileUrl, post } from "@/lib/api";
import type { useSlides } from "./use-slides";

type Zone = PbiZone & { ref?: Ref; pageName?: string; visualName?: string; error?: string };

/** One live Power BI view: a visual alone, or a report page without its panes, with the zone's filters. */
function LiveZone({ zone, scale }: { zone: Zone; scale: number }) {
  const el = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const div = el.current;
    if (!div || !zone.ref) return;
    let gone = false;
    void (async () => {
      const { pbi, service } = await powerbi();
      if (gone) return;
      const filters = zone.filters.map((f) => ({ $schema: "http://powerbi.com/product/schema#basic", target: { table: f.table, column: f.column }, operator: "In", values: f.values, filterType: pbi.models.FilterType.Basic }));
      const common = { id: zone.ref!.id, embedUrl: embedUrl(zone.ref!), accessToken: await embedToken(), tokenType: pbi.models.TokenType.Aad, filters: filters as never, settings: { panes: { filters: { visible: false }, pageNavigation: { visible: false } }, background: pbi.models.BackgroundType.Transparent } };
      service.embed(div, zone.visualName && zone.pageName ? ({ ...common, type: "visual", pageName: zone.pageName, visualName: zone.visualName } as never) : ({ ...common, type: "report", pageName: zone.pageName, settings: { ...common.settings, layoutType: pbi.models.LayoutType.Custom, customLayout: { displayOption: pbi.models.DisplayOption.FitToPage } } } as never));
    })();
    return () => {
      gone = true;
      void powerbi().then(({ service }) => service.reset(div));
    };
  }, [zone.ref?.id, zone.pageName, zone.visualName, JSON.stringify(zone.filters)]); // eslint-disable-line react-hooks/exhaustive-deps
  const b = zone.box;
  return <div ref={el} className="absolute overflow-hidden rounded-md bg-white [&_iframe]:border-0" style={{ left: b.x * scale, top: b.y * scale, width: b.w * scale, height: b.h * scale }} />;
}

function FullScreen({ pics, start, zones, onClose }: { pics: string[]; start: number; zones: Zone[]; onClose: () => void }) {
  const [i, setI] = useState(start);
  const [box, setBox] = useState({ w: 0, h: 0 });
  const root = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = root.current!;
    void el.requestFullscreen?.().catch(() => undefined);
    const key = (e: KeyboardEvent) => {
      if (e.key === "ArrowRight" || e.key === " " || e.key === "PageDown") setI((x) => Math.min(pics.length - 1, x + 1));
      if (e.key === "ArrowLeft" || e.key === "PageUp") setI((x) => Math.max(0, x - 1));
      if (e.key === "Escape") onClose();
    };
    const size = () => setBox({ w: window.innerWidth, h: window.innerHeight });
    const exit = () => !document.fullscreenElement && onClose();
    size();
    window.addEventListener("keydown", key);
    window.addEventListener("resize", size);
    document.addEventListener("fullscreenchange", exit);
    return () => {
      window.removeEventListener("keydown", key);
      window.removeEventListener("resize", size);
      document.removeEventListener("fullscreenchange", exit);
      if (document.fullscreenElement) void document.exitFullscreen().catch(() => undefined);
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  // The slide as large as the screen allows, 16:9.
  const w = Math.min(box.w, (box.h * SLIDE.w) / SLIDE.h);
  const h = (w * SLIDE.h) / SLIDE.w;
  const scale = w / SLIDE.w;
  return (
    <div ref={root} className="fixed inset-0 z-50 flex items-center justify-center bg-black">
      <div className="relative" style={{ width: w, height: h }}>
        <img src={fileUrl(pics[i])} alt={`Slide ${i + 1}`} className="absolute inset-0 h-full w-full" />
        {zones.filter((z) => z.slide === i && z.ref).map((z, k) => <LiveZone key={`${i}-${k}`} zone={z} scale={scale} />)}
      </div>
      <div className="absolute right-4 bottom-4 flex items-center gap-1 rounded-full bg-white/10 px-2 py-1 text-white opacity-30 transition-opacity hover:opacity-100">
        <button onClick={() => setI(Math.max(0, i - 1))} className="p-1.5" aria-label="Previous"><ChevronLeftIcon className="size-5" /></button>
        <span className="text-[13px] tabular-nums">{i + 1} / {pics.length}</span>
        <button onClick={() => setI(Math.min(pics.length - 1, i + 1))} className="p-1.5" aria-label="Next"><ChevronRightIcon className="size-5" /></button>
        <button onClick={onClose} className="p-1.5" aria-label="Exit"><XIcon className="size-5" /></button>
      </div>
    </div>
  );
}

export function Present({ slides, markdown }: { slides: ReturnType<typeof useSlides>; markdown?: string }) {
  const [playing, setPlaying] = useState<number>();
  const [zones, setZones] = useState<Zone[]>([]);
  useEffect(() => {
    if (markdown === undefined) return setZones([]);
    post<Zone[]>("/deck/zones", { markdown }).then(setZones, () => setZones([]));
  }, [markdown]);
  const pics = slides.pics;
  return (
    <div className="flex flex-col gap-4 p-6">
      <div className="flex items-center gap-2">
        <div className="flex-1 text-[13px] text-muted-foreground">{pics ? `${pics.length} slides${zones.some((z) => z.ref) ? " · Power BI views live in full screen" : ""}` : "Drawing the slides…"}</div>
        <Button variant="outline" size="sm" onClick={slides.render} disabled={slides.busy}>{slides.busy ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh</Button>
        <Button variant="dark" size="sm" disabled={!pics?.length} onClick={() => setPlaying(0)}><PlayIcon /> Present</Button>
      </div>
      {slides.error && <ErrorBox>{slides.error}</ErrorBox>}
      <div className="grid gap-4 sm:grid-cols-2 2xl:grid-cols-3">
        {pics?.map((id, i) => (
          <button key={id} onClick={() => setPlaying(i)} className="group text-left" title="Present from this slide">
            <img src={fileUrl(id)} alt={`Slide ${i + 1}`} className="w-full rounded-xl border shadow-[0_1px_4px_rgba(0,0,0,0.06)] transition-transform group-hover:-translate-y-0.5" />
            <div className="mt-1 text-[12px] text-muted-foreground">{i + 1}</div>
          </button>
        ))}
        {!pics && slides.busy && Array.from({ length: 4 }, (_, i) => <div key={i} className="aspect-video animate-pulse rounded-xl bg-muted" />)}
      </div>
      {playing !== undefined && pics && <FullScreen pics={pics} start={playing} zones={zones} onClose={() => setPlaying(undefined)} />}
    </div>
  );
}
