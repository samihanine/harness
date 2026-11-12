/**
 * A React element as text, drawn by the browser's own renderer (react-dom/client) into a detached
 * element: the page never loads the server renderer (react-dom/server). Used where markup has to
 * leave the app as text — the Confluence storage format, the SVG of a diagram. Never call it while
 * React is rendering (flushSync): call it from event handlers and async code.
 */
import type { ReactElement } from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";

function draw(element: ReactElement) {
  const host = document.createElement("div");
  const root = createRoot(host);
  flushSync(() => root.render(element));
  return { host, done: () => root.unmount() };
}

/** The element as HTML text (what `innerHTML` gives: right for SVG). */
export function renderHtml(element: ReactElement) {
  const { host, done } = draw(element);
  const html = host.innerHTML;
  done();
  return html;
}

/** The element as well-formed XHTML (void tags closed: `<br />`), without the namespace the serializer adds. */
export function renderXhtml(element: ReactElement) {
  const { host, done } = draw(element);
  const serializer = new XMLSerializer();
  const xml = Array.from(host.childNodes, (n) => serializer.serializeToString(n)).join("");
  done();
  return xml.replace(/ xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, "");
}
