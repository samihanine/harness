/**
 * A PDF drawn page by page in the browser (pdf.js, its worker inlined: no network, no eval): the
 * pages of the deck as Microsoft Office drew them, as pictures for the slide rail and the slide shown.
 * Loaded when a deck is opened (this file is part of the Presenter's chunk only).
 */
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import PdfWorker from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?worker&inline";

let ready = false;
function setup() {
  if (ready) return;
  pdfjs.GlobalWorkerOptions.workerPort = new PdfWorker();
  ready = true;
}

export type Pdf = pdfjs.PDFDocumentProxy;

export async function openPdf(data: Uint8Array): Promise<Pdf> {
  setup();
  // A copy: pdf.js takes the buffer over.
  return pdfjs.getDocument({ data: data.slice(), isEvalSupported: false, enableXfa: false }).promise;
}

/** One page (1-based) as a PNG data URL, `width` pixels wide (twice that on a dense screen, up to 2400). */
export async function pageImage(pdf: Pdf, number: number, width: number): Promise<string> {
  const page = await pdf.getPage(number);
  const base = page.getViewport({ scale: 1 });
  const scale = Math.min(2400, width * Math.min(2, window.devicePixelRatio || 1)) / base.width;
  const viewport = page.getViewport({ scale });
  const canvas = Object.assign(document.createElement("canvas"), { width: Math.ceil(viewport.width), height: Math.ceil(viewport.height) });
  await page.render({ canvas, canvasContext: canvas.getContext("2d")!, viewport }).promise;
  page.cleanup();
  return canvas.toDataURL("image/png");
}
