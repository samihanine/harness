/**
 * Pictures of the Power BI views of a deck, for its PDF and PowerPoint: a Power BI report cannot be
 * drawn into an image by the page (cross-origin frames), so the browser tab is captured (the user
 * shares it once: getDisplayMedia) while each slide with Power BI is shown, and every live zone on
 * screen is cut from the frame. Works on Power BI Pro (no Export To File API).
 */
import { liveRendered, liveZones, setCapture } from "./powerbi";
import type { Deck } from ".";

const POWER_BI = ["report", "visual", "compare", "tutorial"];
const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** The slides of a deck that show Power BI. */
export const powerBiSlides = (deck: Deck) => deck.slides.flatMap((s, i) => (POWER_BI.includes(s.template) ? [i] : []));

/**
 * Shows each Power BI slide (`show`), waits for its report to draw, and keeps a picture of each zone.
 * Returns how many were captured; throws when the user declines the tab share.
 */
export async function capturePowerBi(deck: Deck, show: (index: number) => void) {
  const indexes = powerBiSlides(deck);
  if (!indexes.length) return 0;
  const stream = await navigator.mediaDevices.getDisplayMedia({ video: { displaySurface: "browser" }, audio: false, preferCurrentTab: true, selfBrowserSurface: "include" } as DisplayMediaStreamOptions);
  const video = Object.assign(document.createElement("video"), { muted: true, playsInline: true, srcObject: stream });
  let count = 0;
  try {
    await video.play();
    for (const index of indexes) {
      show(index);
      // The slide's entrance, then its reports drawing.
      await wait(1200);
      await Promise.race([liveRendered(), wait(20_000)]);
      await wait(1500);
      // The captured frame is the tab's viewport, maybe at another scale (device pixels).
      const kx = video.videoWidth / window.innerWidth;
      const ky = video.videoHeight / window.innerHeight;
      for (const { key, entry } of liveZones.values()) {
        // Each view whole and on top (a compare slider cuts one and lays it over the other).
        entry.whole = true;
        await wait(400);
        const r = entry.host.getBoundingClientRect();
        if (r.width >= 4 && r.height >= 4) {
          const frame = Object.assign(document.createElement("canvas"), { width: video.videoWidth, height: video.videoHeight });
          frame.getContext("2d")!.drawImage(video, 0, 0);
          const cut = Object.assign(document.createElement("canvas"), { width: Math.round(r.width * kx), height: Math.round(r.height * ky) });
          cut.getContext("2d")!.drawImage(frame, r.left * kx, r.top * ky, r.width * kx, r.height * ky, 0, 0, cut.width, cut.height);
          setCapture(key, { url: cut.toDataURL("image/png"), width: cut.width, height: cut.height });
          count++;
        }
        entry.whole = false;
      }
    }
  } finally {
    stream.getTracks().forEach((t) => t.stop());
  }
  return count;
}
