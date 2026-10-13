/**
 * One visual of the embedded report, alone, with any licence: the page is shown with that visual only
 * (custom layout; the hidden slicers and filters still apply), then either printed (PDF through the
 * browser print dialog) or captured from the tab (PNG through screen sharing), then the normal layout
 * comes back.
 */
import { models, type Report } from "./pbi";

/** Shows the visual alone over the whole page; returns how to give the page back as it was. */
async function isolate(report: Report, page: string, visual: string) {
  const m = await models();
  const target = (await report.getPages()).find((p) => p.name === page);
  if (!target) throw new Error("Page not found.");
  if (!target.isActive) await target.setActive();
  const visuals = await target.getVisuals();
  if (!visuals.some((x) => x.name === visual)) throw new Error("Visual not found on the page.");
  // Every visual's own place and visibility, given back as is afterwards (an empty layout does not
  // reset a visual that was moved, and visuals hidden in the report must stay hidden).
  const original = Object.fromEntries(
    visuals.map(({ name, layout: { x = 0, y = 0, z = 0, width = 0, height = 0, displayState } }) => [
      name,
      { x, y, z, width, height, displayState: displayState ?? { mode: m.VisualContainerDisplayMode.Visible } },
    ]),
  );
  // The visual over the whole page (no empty space around it): Power BI redraws it at that size
  // (wider columns, larger chart), it is not stretched like an image.
  const { width: pw = 1280, height: ph = 720 } = target.defaultSize ?? {};
  const layout = (pagesLayout: Record<string, unknown>) =>
    report.updateSettings({ layoutType: m.LayoutType.Custom, customLayout: { displayOption: m.DisplayOption.FitToPage, pagesLayout } } as never);
  await layout({
    [page]: {
      defaultLayout: { displayState: { mode: m.VisualContainerDisplayMode.Hidden } },
      visualsLayout: { [visual]: { x: 0, y: 0, z: 1, width: pw, height: ph, displayState: { mode: m.VisualContainerDisplayMode.Visible } } },
    },
  });
  // Lets the visual draw at its new size.
  await new Promise((r) => setTimeout(r, 1000));
  // Back to the viewer's default view (fit to width).
  return () =>
    report.updateSettings({
      layoutType: m.LayoutType.Custom,
      customLayout: { displayOption: m.DisplayOption.FitToWidth, pagesLayout: { [page]: { defaultLayout: { displayState: { mode: m.VisualContainerDisplayMode.Visible } }, visualsLayout: original } } },
    } as never);
}

export async function printVisual(report: Report, page: string, visual: string) {
  const restore = await isolate(report, page, visual);
  try {
    // Some browsers never settle print(): the normal layout comes back after 30 s at the latest
    // (the print dialog works on its own copy of the page by then).
    await Promise.race([report.print(), new Promise((r) => setTimeout(r, 30_000))]);
  } finally {
    await restore();
  }
}

/* ------------------------------- PNG capture ------------------------------- */

/** Sharing of this tab, kept for the session: only the first capture asks the user. */
let shared: { stream: MediaStream; video: HTMLVideoElement } | undefined;

async function tabStream() {
  if (shared?.stream.getVideoTracks()[0]?.readyState === "live") return shared;
  if (!navigator.mediaDevices?.getDisplayMedia) throw new Error("This browser cannot capture the tab (use Chrome or Edge).");
  const stream = await navigator.mediaDevices
    .getDisplayMedia({ video: { displaySurface: "browser" }, audio: false, preferCurrentTab: true, selfBrowserSurface: "include" } as DisplayMediaStreamOptions)
    .catch((e) => {
      throw new Error(e?.name === "NotAllowedError" ? "Tab sharing refused: the PNG needs it (share this tab)." : String(e?.message ?? e));
    });
  const video = Object.assign(document.createElement("video"), { srcObject: stream, muted: true, playsInline: true });
  // In the page (invisible): a detached video may be throttled and keep showing an old frame.
  video.style.cssText = "position:fixed;left:0;top:0;width:2px;height:2px;opacity:0;pointer-events:none";
  document.body.append(video);
  await video.play();
  stream.getVideoTracks()[0].addEventListener("ended", () => video.remove());
  return (shared = { stream, video });
}

/**
 * The visual as a PNG, captured from the tab: the report area only (cut out by its position in the
 * window), with the even margins around the visual trimmed.
 * Call it from the click itself: the browser only asks for sharing during a user gesture.
 */
export async function captureVisualPng(report: Report, page: string, visual: string, title: string) {
  // Asked first, while the click still counts as a user gesture.
  const { video } = await tabStream();
  const restore = await isolate(report, page, visual);
  try {
    save(await grab(video, (report as unknown as { element: HTMLElement }).element), title || visual);
  } finally {
    await restore();
  }
}

/**
 * PNG of the guide's preview (the visual alone, with the report's filters): for the capture it is shown
 * at its real size over the page (only its style changes: moving it would reload its frame), then put
 * back. The main report is not touched.
 */
export async function capturePreviewPng(element: HTMLElement, title: string) {
  const { video } = await tabStream();
  const before = element.style.cssText;
  const width = element.offsetWidth;
  const height = element.offsetHeight;
  // As large as the window allows, never larger than in the report.
  const scale = Math.min(1, (window.innerWidth - 40) / width, (window.innerHeight - 40) / height);
  Object.assign(element.style, { position: "fixed", left: "20px", top: "20px", zIndex: "9999", background: "#fff", transform: `scale(${scale})`, opacity: "1" });
  try {
    await new Promise((r) => setTimeout(r, 700)); // redrawn at that size, and a fresh frame sent
    save(await grab(video, element), title);
  } finally {
    element.style.cssText = before;
  }
}

/** The tab's current frame, cut to an element and trimmed of its even margins, as a PNG data URL. */
async function grab(video: HTMLVideoElement, element: HTMLElement) {
  // The whole tab is captured and the area cut out by its position. (Region Capture was dropped: the
  // tab only sends a new frame when the page changes, so after a crop the video kept showing the
  // previous, uncropped frame, and the second PNG was the whole screen.)
  await new Promise((r) => setTimeout(r, 500));
  const frame = document.createElement("canvas");
  frame.width = video.videoWidth;
  frame.height = video.videoHeight;
  const g = frame.getContext("2d", { willReadFrequently: true })!;
  g.drawImage(video, 0, 0);
  const box = trim(g, areaOf(element, frame));
  const out = document.createElement("canvas");
  out.width = box.w;
  out.height = box.h;
  out.getContext("2d")!.drawImage(frame, box.x, box.y, box.w, box.h, 0, 0, box.w, box.h);
  return out.toDataURL("image/png");
}

function save(href: string, title: string) {
  const a = document.createElement("a");
  a.href = href;
  a.download = `${title.replace(/[^\w\- ]+/g, "").trim() || "visual"}.png`;
  a.click();
}

type Box = { x: number; y: number; w: number; h: number };

/** Place of an element in a frame of the whole tab (the frame is the viewport, maybe at another scale). */
function areaOf(element: HTMLElement, frame: HTMLCanvasElement): Box {
  const r = element.getBoundingClientRect();
  const sx = frame.width / window.innerWidth;
  const sy = frame.height / window.innerHeight;
  return { x: Math.round(r.left * sx), y: Math.round(r.top * sy), w: Math.round(r.width * sx), h: Math.round(r.height * sy) };
}

/** The area without its even margins (same color as its top left corner), with a small padding kept. */
function trim(g: CanvasRenderingContext2D, area: Box, padding = 12): Box {
  const { data } = g.getImageData(area.x, area.y, area.w, area.h);
  const [r0, g0, b0] = [data[0], data[1], data[2]];
  let top = area.h;
  let left = area.w;
  let bottom = -1;
  let right = -1;
  for (let y = 0; y < area.h; y++)
    for (let x = 0; x < area.w; x++) {
      const i = (y * area.w + x) * 4;
      if (Math.abs(data[i] - r0) + Math.abs(data[i + 1] - g0) + Math.abs(data[i + 2] - b0) > 24) {
        if (y < top) top = y;
        if (y > bottom) bottom = y;
        if (x < left) left = x;
        if (x > right) right = x;
      }
    }
  if (bottom < 0) return area; // nothing drawn: kept whole
  const x = Math.max(0, left - padding);
  const y = Math.max(0, top - padding);
  return { x: area.x + x, y: area.y + y, w: Math.min(area.w, right + padding + 1) - x, h: Math.min(area.h, bottom + padding + 1) - y };
}
