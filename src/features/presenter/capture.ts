/**
 * Pictures of Power BI visuals for the slides. A Power BI report cannot be drawn into an image by the
 * page (cross-origin frames) and the Export To File API needs a Premium capacity, so the visual is
 * embedded on screen for a moment and the browser tab is captured (the user shares it once per
 * session: getDisplayMedia), then cut to the visual. Works on Power BI Pro.
 */
import { embedReport, embedVisual, resetEmbed, settled, type Ref } from "@/services/pbi";

let stream: MediaStream | undefined;
let video: HTMLVideoElement | undefined;
const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** The tab's picture stream, asked once (kept while the tab is shared). */
async function frames() {
  if (stream?.active && video) return video;
  stream = await navigator.mediaDevices.getDisplayMedia({ video: { displaySurface: "browser" }, audio: false, preferCurrentTab: true, selfBrowserSurface: "include" } as DisplayMediaStreamOptions).catch((e) => {
    throw new Error(`The tab was not shared, so the visual cannot be captured (${e instanceof Error ? e.message : e}).`);
  });
  video = Object.assign(document.createElement("video"), { muted: true, playsInline: true, srcObject: stream });
  await video.play();
  stream.getVideoTracks()[0]?.addEventListener("ended", () => (stream = undefined));
  return video;
}

/** Stops sharing the tab (when the screen is left). */
export const stopCapture = () => (stream?.getTracks().forEach((t) => t.stop()), (stream = undefined));

/**
 * A visual (or a whole page) of a report as a PNG: embedded over the screen at its proportions
 * (`aspect`: width / height), drawn, captured, cut.
 */
export async function captureVisual(report: Ref, page: string, visual?: string, aspect = 16 / 9): Promise<{ png: Uint8Array; width: number; height: number }> {
  const v = await frames();
  const width = Math.min(1200, window.innerWidth - 80, (window.innerHeight - 120) * aspect);
  const height = width / aspect;
  const shade = Object.assign(document.createElement("div"), { style: "position:fixed;inset:0;z-index:2147483646;background:rgba(15,23,42,.55);display:flex;align-items:center;justify-content:center" });
  const host = Object.assign(document.createElement("div"), { style: `width:${width}px;height:${height}px;background:#fff;border-radius:8px;overflow:hidden` });
  shade.appendChild(host);
  document.body.appendChild(shade);
  try {
    const embed = visual ? await embedVisual(host, report, page, visual) : await embedReport(host, report, "view", { bare: true });
    if (!visual) await (embed as Awaited<ReturnType<typeof embedReport>>).setPage(page).catch(() => undefined);
    await settled(embed as never);
    await wait(1800);
    const r = host.getBoundingClientRect();
    const kx = v.videoWidth / window.innerWidth;
    const ky = v.videoHeight / window.innerHeight;
    const frame = Object.assign(document.createElement("canvas"), { width: v.videoWidth, height: v.videoHeight });
    frame.getContext("2d")!.drawImage(v, 0, 0);
    const cut = Object.assign(document.createElement("canvas"), { width: Math.round(r.width * kx), height: Math.round(r.height * ky) });
    cut.getContext("2d")!.drawImage(frame, r.left * kx, r.top * ky, r.width * kx, r.height * ky, 0, 0, cut.width, cut.height);
    const blob = await new Promise<Blob>((resolve, reject) => cut.toBlob((b) => (b ? resolve(b) : reject(new Error("The capture failed."))), "image/png"));
    return { png: new Uint8Array(await blob.arrayBuffer()), width: cut.width, height: cut.height };
  } finally {
    await resetEmbed(host).catch(() => undefined);
    shade.remove();
  }
}
