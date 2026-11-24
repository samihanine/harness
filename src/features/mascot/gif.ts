/**
 * The mascot's animations as GIF files (Settings › General): the animated SVG is drawn frame by frame
 * (its current transforms serialized, drawn on a canvas) and encoded with gifenc.
 */
import { GIFEncoder, applyPalette, quantize } from "gifenc";

export async function recordGif(svgOf: () => SVGSVGElement | null, { seconds = 3, fps = 20, size = 240 } = {}) {
  const gif = GIFEncoder();
  const canvas = Object.assign(document.createElement("canvas"), { width: size, height: size });
  const ctx = canvas.getContext("2d", { willReadFrequently: true })!;
  const frames = Math.round(seconds * fps);
  for (let i = 0; i < frames; i++) {
    const svg = svgOf();
    if (!svg) break;
    const clone = svg.cloneNode(true) as SVGSVGElement;
    clone.setAttribute("width", String(size));
    clone.setAttribute("height", String(size));
    const url = URL.createObjectURL(new Blob([new XMLSerializer().serializeToString(clone)], { type: "image/svg+xml" }));
    const img = new Image();
    await new Promise((res, rej) => ((img.onload = res), (img.onerror = rej), (img.src = url)));
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, size, size);
    const k = Math.min(size / img.width, size / img.height);
    ctx.drawImage(img, (size - img.width * k) / 2, (size - img.height * k) / 2, img.width * k, img.height * k);
    URL.revokeObjectURL(url);
    const { data } = ctx.getImageData(0, 0, size, size);
    const palette = quantize(data, 256);
    gif.writeFrame(applyPalette(data, palette), size, size, { palette, delay: 1000 / fps });
    await new Promise((r) => setTimeout(r, 1000 / fps));
  }
  gif.finish();
  return new Blob([gif.bytes() as BlobPart], { type: "image/gif" });
}
