/**
 * React drawings (the mascot, an icon, the title pages' flowers) as PNG pictures for PowerPoint: drawn
 * off screen, the SVG made standalone (the styles CSS gives it written on each element), then painted
 * on a canvas.
 */
import { type ReactElement } from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";

/** Styles the CSS may give an SVG's elements, written on each so the SVG stands alone. */
const INLINE = ["fill", "stroke", "stroke-width", "stroke-opacity", "fill-opacity", "opacity", "stroke-linecap", "stroke-linejoin", "display", "vector-effect", "color"];

/**
 * The first SVG `element` draws, as PNG data (for PowerPoint) with its size. `host`: the class and
 * style of the box it is drawn in (e.g. the deck's theme variables, "slide-still"); `scale`: pixels
 * per SVG unit. Undefined when it cannot be drawn.
 */
export async function svgPicture(element: ReactElement, { width, height, scale = 2, className = "", style = {} }: { width: number; height: number; scale?: number; className?: string; style?: Record<string, string> }) {
  const host = Object.assign(document.createElement("div"), { className });
  for (const [k, v] of Object.entries({ ...style, position: "fixed", left: "-10000px", top: "0", width: `${width}px`, height: `${height}px` })) host.style.setProperty(k.startsWith("--") ? k : k.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`), v);
  document.body.append(host);
  const root = createRoot(host);
  try {
    flushSync(() => root.render(element));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    const svg = host.querySelector("svg");
    if (!svg) return undefined;
    const copy = svg.cloneNode(true) as SVGSVGElement;
    const originals = [svg, ...svg.querySelectorAll("*")];
    [copy, ...copy.querySelectorAll("*")].forEach((el, i) => {
      const css = getComputedStyle(originals[i]);
      el.setAttribute("style", `${INLINE.map((p) => `${p}:${css.getPropertyValue(p)}`).join(";")};stroke-dasharray:none;animation:none`);
      el.removeAttribute("class");
    });
    copy.setAttribute("width", String(width));
    copy.setAttribute("height", String(height));
    copy.setAttribute("xmlns", "http://www.w3.org/2000/svg");
    const img = new Image();
    await new Promise((resolve, reject) => ((img.onload = resolve), (img.onerror = reject), (img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(new XMLSerializer().serializeToString(copy))}`)));
    const canvas = Object.assign(document.createElement("canvas"), { width: Math.round(width * scale), height: Math.round(height * scale) });
    canvas.getContext("2d")!.drawImage(img, 0, 0, canvas.width, canvas.height);
    return { data: canvas.toDataURL("image/png").slice(5), width, height };
  } catch {
    return undefined;
  } finally {
    root.unmount();
    host.remove();
  }
}
