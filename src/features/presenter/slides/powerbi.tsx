/**
 * Power BI in the slides. Each report page is embedded once (Power BI JavaScript SDK, with the app's
 * token: no sign-in inside the slide) and laid over the slide shown: every slide of that page (the
 * page, one visual by its zone, a tutorial) reuses the same embed, moved and scaled over its box. Only
 * the slide on screen is live; thumbnails, the PDF and the PowerPoint show a capture of that zone
 * (taken from the tab, captures below), or a placeholder until there is one.
 */
import { createContext, useContext, useEffect, useRef, useState, useSyncExternalStore, type CSSProperties, type ReactNode } from "react";
import { createPortal } from "react-dom";
import { z } from "zod";
import { PERFORMANCE } from "@/config";
import { embedReport, models, resetEmbed, type Report } from "@/services/pbi";
import { all, useCollection } from "@/lib/store";
import type PptxGenJS from "pptxgenjs";
import { hex, type Theme } from "./kit";

export const PAGE = { w: 1280, h: 720 };
export const cropSchema = z.object({ x: z.number(), y: z.number(), w: z.number(), h: z.number() }).describe("Zone of the page in px of a 1280×720 page");
export type Crop = z.infer<typeof cropSchema>;

/** Set by the slide view: a still slide (thumbnail, print) shows captures, never a live report. */
export const StillSlide = createContext(false);

/** The report of the app given by its id or its name. */
export function useReportRef(report: string) {
  const reports = useCollection("reports");
  return reports.find((r) => r.id === report) ?? reports.find((r) => r.name.toLowerCase() === report.toLowerCase());
}

export const embedUrl = (r: { id: string; groupId?: string }, page?: string) =>
  `https://app.powerbi.com/reportEmbed?reportId=${r.id}${r.groupId ? `&groupId=${r.groupId}` : ""}&autoAuth=true&navContentPaneEnabled=false&filterPaneEnabled=false${page ? `&pageName=${encodeURIComponent(page)}` : ""}`;

/** A visual's zone on its page (1280×720 px), from the schema kept with the report (Reports screen). */
export function visualZone(entry: ReturnType<typeof useReportRef>, page: string | undefined, visual: string): Crop | undefined {
  const schema = entry?.schema;
  const v = schema?.visuals.find((x) => x.name === visual && (!page || x.page === page)) ?? schema?.visuals.find((x) => x.title === visual);
  if (!v) return undefined;
  const p = schema!.pages.find((x) => x.name === v.page);
  const k = 1280 / (p?.width || 1280);
  return { x: v.x * k, y: v.y * k, w: v.width * k, h: v.height * k };
}

/** The smallest zone holding all the given zones (a frame around several visuals). */
export function unionZone(zones: (Crop | undefined)[]): Crop | undefined {
  const list = zones.filter((z): z is Crop => !!z);
  if (!list.length) return undefined;
  const x = Math.min(...list.map((z) => z.x));
  const y = Math.min(...list.map((z) => z.y));
  return { x, y, w: Math.max(...list.map((z) => z.x + z.w)) - x, h: Math.max(...list.map((z) => z.y + z.h)) - y };
}

/* -------------------------------- Captures -------------------------------- */

/** The key of a Power BI zone of a slide: its report, page and crop (the same zone, the same picture). */
export const captureKey = (report: string, page?: string, crop?: Crop) => `${report}|${page ?? ""}|${crop ? `${Math.round(crop.x)},${Math.round(crop.y)},${Math.round(crop.w)},${Math.round(crop.h)}` : ""}`;
const captures = new Map<string, { url: string; width: number; height: number }>();
const listeners = new Set<() => void>();
let version = 0;
export const getCapture = (key: string) => captures.get(key);
export const setCapture = (key: string, picture: { url: string; width: number; height: number }) => (captures.set(key, picture), version++, listeners.forEach((l) => l()));
const useCaptures = () =>
  useSyncExternalStore(
    (l) => (listeners.add(l), () => void listeners.delete(l)),
    () => version,
  );

/** The live Power BI zones on screen now (key and box): what an export captures. */
export const liveZones = new Map<HTMLElement, { key: string; entry: Entry }>();

/* -------------------------------- PowerPoint -------------------------------- */

/** The app's reports for the PowerPoint export, by id and by name (filled before it runs: the export is not a component). */
const reportsByRef = new Map<string, { id: string; url: string }>();
export async function loadReportLinks() {
  for (const r of await all("reports")) for (const k of [r.id, r.name]) reportsByRef.set(k, { id: r.id, url: embedUrl(r) });
}

/**
 * A Power BI zone in a PowerPoint slide (inches): the capture taken from the screen (contained in the
 * box, linking to the live report), or a framed placeholder with the link when there is none.
 */
export function pptxPowerBi(slide: PptxGenJS.Slide, theme: Theme, view: { report: string; page?: string; crop?: Crop }, box: { x: number; y: number; w: number; h: number }) {
  const ref = reportsByRef.get(view.report);
  const url = ref?.url ?? "https://app.powerbi.com";
  const picture = getCapture(captureKey(ref?.id ?? view.report, view.page, view.crop));
  if (picture) {
    const ratio = picture.width / picture.height;
    const w = Math.min(box.w, box.h * ratio);
    const h = w / ratio;
    slide.addImage({ data: picture.url.slice(5), x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h, hyperlink: { url } });
    return;
  }
  slide.addShape("roundRect", { ...box, fill: { color: hex(theme.primary), transparency: 94 }, line: { color: hex(theme.primary), transparency: 60, dashType: "dash" }, rectRadius: 0.1 });
  slide.addText([{ text: "Power BI", options: { bold: true, fontSize: 18, breakLine: true } }, { text: "Open the live report", options: { fontSize: 13, hyperlink: { url } } }], { ...box, align: "center", valign: "middle", color: hex(theme.primary), fontFace: theme.fontText });
}

/* ---------------------------------- Embeds --------------------------------- */

/** `height`: the page's height at 1280 px wide (pages may be taller than 16:9); `whole`: shown uncut on top (captured). */
type Entry = { base: string; host: HTMLDivElement; clip: HTMLDivElement; stage: HTMLDivElement; frames: HTMLDivElement; report?: Report; height: number; rendered: Promise<void>; error?: string; users: number; used: number; whole?: boolean };
const pool: Entry[] = [];

/** The layer over the app holding the embeds (fixed, above the slides). */
function layer() {
  let el = document.getElementById("pbi-layer");
  if (!el) {
    el = Object.assign(document.createElement("div"), { id: "pbi-layer" });
    Object.assign(el.style, { position: "fixed", inset: "0", pointerEvents: "none", zIndex: "30" });
    document.body.append(el);
    // In fullscreen only the fullscreen element shows: the layer goes into it, and back after.
    const host = el;
    document.addEventListener("fullscreenchange", () => (document.fullscreenElement ?? document.body).append(host));
  }
  return el;
}

/** An embed of a report page for one slide zone: an idle one of that page if there is one, else a new one (the oldest idle ones beyond the limit are dropped). */
function acquire(ref: { id: string; groupId?: string }, page?: string): Entry {
  const base = `${ref.id}|${page ?? ""}`;
  let entry = pool.find((e) => e.base === base && e.users === 0);
  if (!entry) {
    const host = document.createElement("div");
    // Hidden off-screen, never "display: none": a report loading at 0×0 stays blank.
    Object.assign(host.style, { position: "fixed", left: "-99999px", top: "0", width: `${PAGE.w}px`, height: `${PAGE.h}px`, pointerEvents: "auto" });
    const clip = document.createElement("div");
    Object.assign(clip.style, { position: "absolute", overflow: "hidden", borderRadius: "8px", background: "#fff", width: `${PAGE.w}px`, height: `${PAGE.h}px` });
    const stage = document.createElement("div");
    Object.assign(stage.style, { position: "absolute", left: "0", top: "0", width: `${PAGE.w}px`, height: `${PAGE.h}px`, transformOrigin: "0 0" });
    const frames = document.createElement("div");
    Object.assign(frames.style, { position: "absolute", inset: "0", pointerEvents: "none" });
    clip.append(stage);
    host.append(clip, frames);
    layer().append(host);
    const e: Entry = { base, host, clip, stage, frames, height: PAGE.h, users: 0, used: Date.now(), rendered: Promise.resolve() };
    e.rendered = embedReport(stage, ref, "view", { bare: true })
      .then(async (report) => {
        e.report = report;
        if (page) await report.setPage(page).catch(() => undefined);
        // The page drawn at 1280 px wide and its own height: crops are in that page's px.
        const active = await report.getActivePage().catch(() => undefined);
        const size = active?.defaultSize;
        if (size?.width && size.height) e.height = Math.round((PAGE.w * size.height) / size.width);
        e.stage.style.height = `${e.height}px`;
        const m = await models();
        await report.updateSettings({ layoutType: m.LayoutType.Custom, customLayout: { displayOption: m.DisplayOption.FitToWidth } }).catch(() => undefined);
        await new Promise<void>((resolve) => {
          const done = () => (report.off("rendered", done), resolve());
          report.on("rendered", done);
          setTimeout(done, 8000);
        });
      })
      .catch((err) => void (e.error = err instanceof Error ? err.message : String(err)));
    entry = e;
    pool.push(e);
    // Beyond the limit, the idle embeds used longest ago go.
    const idle = pool.filter((x) => x.users === 0 && x !== e).sort((a, b) => a.used - b.used);
    while (pool.length > Math.max(1, PERFORMANCE.maxEmbeds) && idle.length) {
      const old = idle.shift()!;
      void resetEmbed(old.stage).catch(() => undefined);
      old.host.remove();
      pool.splice(pool.indexOf(old), 1);
    }
  }
  entry.users++;
  entry.used = Date.now();
  return entry;
}

const release = (e: Entry) => {
  e.users = Math.max(0, e.users - 1);
  e.used = Date.now();
  if (!e.users) e.host.style.left = "-99999px";
};

/** Waits until the embeds of the zones on screen have drawn their page (export captures). */
export const liveRendered = () => Promise.all([...liveZones.values()].map((z) => z.entry.rendered));

/* ---------------------------------- Frame ---------------------------------- */

type Place = (zone: Crop) => CSSProperties;

/**
 * A report page (or a zone of it, `crop`) in its box: live on the slide shown, the capture elsewhere.
 * `overlay`: drawn over the report in page px (e.g. the tutorial's frames). `reveal`: the share of its
 * width hidden from the left (0–1), showing what is under it (the compare slider).
 */
export function PowerBiFrame({ report, page, crop, overlay, reveal }: { report: string; page?: string; crop?: Crop; overlay?: (place: Place) => ReactNode; reveal?: number }) {
  const still = useContext(StillSlide);
  const ref = useReportRef(report);
  const key = captureKey(ref?.id ?? report, page, crop);
  useCaptures();
  const capture = getCapture(key);
  if (!ref) return <Placeholder text={`Power BI report "${report}" is not in the app (add it in Reports).`} />;
  if (still) return capture ? <img src={capture.url} alt="" className="size-full object-contain" style={reveal ? { clipPath: `inset(0 0 0 ${reveal * 100}%)` } : undefined} /> : <Placeholder text={`Power BI · ${ref.name}`} />;
  return <LiveFrame refId={{ id: ref.id, groupId: ref.groupId }} page={page} crop={crop} captureId={key} overlay={overlay} reveal={reveal} />;
}

function Placeholder({ text }: { text: string }) {
  return (
    <div className="flex size-full items-center justify-center rounded-[0.8cqw] p-[2cqw] text-center text-[1.2cqw] opacity-60" style={{ background: "color-mix(in srgb, var(--s-primary) 6%, var(--s-bg))", border: "1px dashed color-mix(in srgb, var(--s-primary) 40%, transparent)" }}>
      {text}
    </div>
  );
}

/** The live embed laid over this box, following it every frame (the slide may move, scale or scroll). */
function LiveFrame({ refId, page, crop, captureId, overlay, reveal }: { refId: { id: string; groupId?: string }; page?: string; crop?: Crop; captureId: string; overlay?: (place: Place) => ReactNode; reveal?: number }) {
  const box = useRef<HTMLDivElement>(null);
  // Read by the loop below at each frame (the slider moves without starting the embed again).
  const hidden = useRef(reveal);
  hidden.current = reveal;
  const [entry, setEntry] = useState<Entry>();
  const [, setDrawn] = useState(0);
  const scale = useRef(0);
  const zone = crop ?? { x: 0, y: 0, ...PAGE };
  useEffect(() => {
    // Taken and given back by the effect itself (StrictMode runs it twice: never a released embed).
    const e = acquire(refId, page);
    const el = box.current!;
    setEntry(e);
    let live = true;
    // Its error shows once it is known.
    void e.rendered.then(() => live && setDrawn((n) => n + 1));
    liveZones.set(el, { key: captureId, entry: e });
    let frame = 0;
    const follow = () => {
      const r = el.getBoundingClientRect();
      const shown = r.width > 4 && r.height > 4 && el.offsetParent !== null && !e.error;
      if (!shown) e.host.style.left = "-99999px";
      else {
        // A whole page: its own height once known.
        if (!crop) zone.h = e.height;
        const k = Math.min(r.width / zone.w, r.height / zone.h);
        scale.current = k;
        const w = zone.w * k;
        const h = zone.h * k;
        Object.assign(e.host.style, { left: `${r.left + (r.width - w) / 2}px`, top: `${r.top + (r.height - h) / 2}px`, width: `${w}px`, height: `${h}px` });
        Object.assign(e.clip.style, { width: `${w}px`, height: `${h}px` });
        e.stage.style.transform = `scale(${k}) translate(${-zone.x}px, ${-zone.y}px)`;
        // A revealing view is above the other one, cut from the left.
        // `whole`: being captured, whole and above the rest.
        e.clip.style.clipPath = hidden.current === undefined || e.whole ? "" : `inset(0 0 0 ${hidden.current * 100}%)`;
        e.host.style.zIndex = e.whole ? "5" : hidden.current === undefined ? "1" : "2";
      }
      frame = requestAnimationFrame(follow);
    };
    follow();
    return () => {
      live = false;
      cancelAnimationFrame(frame);
      liveZones.delete(el);
      e.clip.style.clipPath = "";
      release(e);
    };
  }, [refId.id, page, zone.x, zone.y, zone.w, zone.h]); // eslint-disable-line react-hooks/exhaustive-deps
  const place: Place = (z) => ({ position: "absolute", left: (z.x - zone.x) * scale.current, top: (z.y - zone.y) * scale.current, width: z.w * scale.current, height: z.h * scale.current });
  return (
    <div ref={box} className="size-full">
      {entry?.error && <Placeholder text={entry.error} />}
      {overlay && entry && createPortal(<OverlayFollow place={place}>{overlay}</OverlayFollow>, entry.frames)}
    </div>
  );
}

/** The overlay drawn again a few times a second (it is placed in screen px and the embed's scale changes). */
function OverlayFollow({ place, children }: { place: Place; children: (place: Place) => ReactNode }) {
  const [, setTick] = useState(0);
  useEffect(() => {
    const t = setInterval(() => setTick((n) => n + 1), 250);
    return () => clearInterval(t);
  }, []);
  return <>{children(place)}</>;
}
