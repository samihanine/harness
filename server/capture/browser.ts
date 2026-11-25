/**
 * "See what the user sees" without screen sharing: a headless browser on the server (the Edge or Chrome
 * already installed, else Playwright's Chromium) embeds the same report with the user's token, applies the
 * state of the user's view (a bookmark state captured in the browser: page, filters, slicers) and takes a
 * picture. Works on Power BI Pro (no Export To File API). The browser and the last reports stay open, so
 * a picture takes a few seconds. The same page renders PDF pages to pictures (pdf.js) for the Presenter.
 */
import type { Browser, BrowserContext, Page } from "playwright-core";
import type { Ref } from "@shared/types";
import { accessToken } from "../ms/auth";
import { staticFile } from "../static";

let browser: Promise<Browser> | undefined;
let context: BrowserContext | undefined;

async function launch(): Promise<Browser> {
  const { chromium } = await import("playwright-core");
  const proxy = process.env.HTTPS_PROXY ?? process.env.https_proxy;
  const options = { headless: true, args: ["--no-sandbox", "--disable-dev-shm-usage"], ...(proxy ? { proxy: { server: proxy, bypass: process.env.NO_PROXY ?? "localhost,127.0.0.1" } } : {}) };
  const attempts = [() => chromium.launch(options), () => chromium.launch({ ...options, channel: "msedge" }), () => chromium.launch({ ...options, channel: "chrome" }), () => chromium.launch({ ...options, executablePath: "/opt/pw-browsers/chromium" })];
  let last: unknown;
  for (const attempt of attempts) {
    try {
      return await attempt();
    } catch (e) {
      last = e;
    }
  }
  throw new Error(`No browser to take pictures with (install Edge or Chrome): ${last instanceof Error ? last.message.split("\n")[0] : last}`);
}

async function newPage(): Promise<Page> {
  browser ??= launch().catch((e) => {
    browser = undefined;
    throw e;
  });
  const b = await browser;
  context ??= await b.newContext({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 1 });
  return context.newPage();
}

const script = async (path: string) => {
  const f = staticFile(`node_modules/${path}`);
  if (!f) throw new Error(`Missing ${path}`);
  return f.text();
};

/* ------------------------------- Power BI ------------------------------- */

type Embedded = { page: Page; ref: Ref; usedAt: number; size: string; state?: string };
const reports = new Map<string, Promise<Embedded>>();
const MAX_OPEN = 3;

const embedUrl = (ref: Ref) => `https://app.powerbi.com/reportEmbed?reportId=${ref.id}${ref.appId ? `&appId=${ref.appId}` : ref.groupId ? `&groupId=${ref.groupId}` : ""}`;

/** The report embedded in a page of the headless browser (kept open for the next pictures). */
async function embedded(ref: Ref, width: number, height: number): Promise<Embedded> {
  const key = ref.id;
  const size = `${width}x${height}`;
  const open = reports.get(key);
  if (open) {
    const e = await open.catch(() => undefined);
    if (e && !e.page.isClosed() && e.size === size) {
      e.usedAt = Date.now();
      return e;
    }
    reports.delete(key);
    await e?.page.close().catch(() => undefined);
  }
  const made = (async () => {
    const page = await newPage();
    await page.setViewportSize({ width, height });
    await page.setContent(`<!doctype html><html><body style="margin:0;background:#fff"><div id="r" style="width:${width}px;height:${height}px"></div></body></html>`);
    await page.addScriptTag({ content: await script("powerbi-client/dist/powerbi.min.js") });
    const token = await accessToken("powerbi");
    await page.evaluate(
      ({ id, url, token }) =>
        new Promise<void>((resolve, reject) => {
          const pbi = (window as any)["powerbi-client"];
          const service = (window as any).powerbi;
          const report = service.embed(document.getElementById("r"), {
            type: "report",
            id,
            embedUrl: url,
            accessToken: token,
            tokenType: pbi.models.TokenType.Aad,
            permissions: pbi.models.Permissions.Read,
            viewMode: pbi.models.ViewMode.View,
            settings: { panes: { filters: { visible: false }, pageNavigation: { visible: false } }, background: pbi.models.BackgroundType.Default, layoutType: pbi.models.LayoutType.Custom, customLayout: { displayOption: pbi.models.DisplayOption.FitToPage } },
          });
          (window as any).report = report;
          const timer = setTimeout(() => reject(new Error("The report did not load in 60 s")), 60_000);
          // Ready when the first page has drawn, not just loaded.
          report.on("rendered", () => (clearTimeout(timer), resolve()));
          report.on("error", (e: any) => (clearTimeout(timer), reject(new Error(e.detail?.detailedMessage ?? e.detail?.message ?? "Could not load the report"))));
        }),
      { id: ref.id, url: embedUrl(ref), token },
    );
    return { page, ref, usedAt: Date.now(), size };
  })();
  reports.set(key, made);
  made.catch(() => reports.delete(key));
  // The least recently used report closed past a few.
  if (reports.size > MAX_OPEN) {
    const all = await Promise.all([...reports.entries()].map(async ([k, p]) => [k, await p.catch(() => undefined)] as const));
    const oldest = all.filter(([, e]) => e).sort((a, b) => a[1]!.usedAt - b[1]!.usedAt)[0];
    if (oldest && oldest[0] !== key) {
      reports.delete(oldest[0]);
      await oldest[1]!.page.close().catch(() => undefined);
    }
  }
  return made;
}

/** Forgets a report open in the headless browser (it was saved: the next picture loads the new version). */
export async function forgetReport(id: string) {
  const open = reports.get(id);
  reports.delete(id);
  const e = await open?.catch(() => undefined);
  await e?.page.close().catch(() => undefined);
}

/** Opens a report in the headless browser ahead of time: the first picture then takes about a second. */
export function warmReport(ref: Ref) {
  void embedded(ref, 1280, 720).catch(() => undefined);
}

export type CaptureOptions = {
  /** The user's view as a bookmark state (page, filters, slicers): what they see. */
  state?: string;
  /** Page name or display name (without state: the report's first page). */
  page?: string;
  /** Filters to apply on the report (Power BI filter objects), e.g. for a deck. */
  filters?: unknown[];
  /** A visual (name or title): only its zone is kept. */
  visual?: string;
  width?: number;
  height?: number;
};

/** A picture of a report page (or of one visual) as the given state shows it, and what is on that page. */
export async function captureReport(ref: Ref, o: CaptureOptions = {}): Promise<{ png: Buffer; page: { name: string; displayName: string }; pages: { name: string; displayName: string }[]; visuals: { name: string; type: string; title?: string }[] }> {
  const width = o.width ?? 1280;
  const height = o.height ?? 720;
  const e = await embedded(ref, width, height);
  const info = await e.page.evaluate(
    async ({ state, keep, page, filters, visual }) => {
      const report = (window as any).report;
      const rendered = () =>
        new Promise<void>((resolve) => {
          const done = () => (report.off("rendered"), resolve());
          report.on("rendered", done);
          // A change that draws nothing new sends no "rendered": not waited for long.
          setTimeout(done, 6_000);
        });
      const pages = await report.getPages();
      let changed = false;
      if (keep) {
        // The view already shown: nothing to apply.
      } else if (state) {
        const wait = rendered();
        await report.bookmarksManager.applyState(state);
        await wait;
        changed = true;
      } else {
        // A clean view: no filter left from a previous picture, the page asked (or the first).
        await report.removeFilters().catch(() => undefined);
        const target = page ? pages.find((p: any) => p.name === page || p.displayName.toLowerCase() === String(page).toLowerCase()) : pages.find((p: any) => p.isActive) ?? pages[0];
        if (page && !target) throw new Error(`No page "${page}". Pages: ${pages.map((p: any) => `"${p.displayName}" (${p.name})`).join(", ")}`);
        const active = pages.find((p: any) => p.isActive);
        if (target && (!active || active.name !== target.name)) {
          const wait = rendered();
          await target.setActive();
          await wait;
          changed = true;
        }
      }
      if (filters?.length) {
        const wait = rendered();
        await report.updateFilters(1, filters).catch(async () => report.setFilters(filters));
        await wait;
        changed = true;
      }
      if (!changed) await new Promise((r) => setTimeout(r, 300));
      const current = (await report.getPages()).find((p: any) => p.isActive);
      const visuals = await current.getVisuals();
      let clip: { x: number; y: number; width: number; height: number } | undefined;
      if (visual) {
        const v = visuals.find((x: any) => x.name === visual) ?? visuals.find((x: any) => (x.title ?? "").toLowerCase() === String(visual).toLowerCase());
        if (!v) throw new Error(`No visual "${visual}" on page "${current.displayName}". Visuals: ${visuals.map((x: any) => `"${x.title ?? x.type}" (${x.name})`).join(", ")}`);
        // The page is drawn fitted to the frame: positions scaled by the page's size.
        const size = current.defaultSize ?? { width: 1280, height: 720 };
        const frame = document.querySelector("#r iframe")!.getBoundingClientRect();
        const k = Math.min(frame.width / size.width, frame.height / size.height);
        const ox = (frame.width - size.width * k) / 2, oy = (frame.height - size.height * k) / 2;
        clip = { x: Math.max(0, ox + v.layout.x * k), y: Math.max(0, oy + v.layout.y * k), width: v.layout.width * k, height: v.layout.height * k };
      }
      return {
        page: { name: current.name, displayName: current.displayName },
        pages: pages.map((p: any) => ({ name: p.name, displayName: p.displayName })),
        visuals: visuals.filter((v: any) => v.layout?.displayState?.mode !== 1).map((v: any) => ({ name: v.name, type: v.type, title: v.title })),
        clip,
      };
    },
    { state: o.state, keep: !!o.state && o.state === e.state, page: o.page, filters: o.filters, visual: o.visual },
  );
  // The same view again (state already applied): nothing to redraw.
  e.state = o.state && !o.filters?.length ? o.state : undefined;
  // Visuals draw a little after "rendered" (animations).
  await e.page.waitForTimeout(800);
  const png = await e.page.screenshot({ type: "png", clip: info.clip ?? { x: 0, y: 0, width, height } });
  return { png, ...info };
}

/** What a hidden embed shows of a report the API cannot read (view only, sensitivity label): pages, visuals, fields, filters. */
export async function embeddedStructure(ref: Ref) {
  const e = await embedded(ref, 1280, 720);
  return e.page.evaluate(async () => {
    const report = (window as any).report;
    const field = (t: any) => (!t ? "?" : t.measure ? `${t.table}[${t.measure}]` : `${t.table}[${t.column ?? t.hierarchyLevel ?? t.hierarchy}]`);
    const pages = await report.getPages();
    const out: any[] = [];
    for (const p of pages) {
      const visuals = await p.getVisuals().catch(() => []);
      const filters = await p.getFilters().catch(() => []);
      out.push({
        name: p.name,
        displayName: p.displayName,
        hidden: p.visibility === 1,
        size: p.defaultSize,
        filters: filters.map((f: any) => ({ field: field(f.target), values: f.values, operator: f.operator })),
        visuals: await Promise.all(
          visuals.map(async (v: any) => {
            const slicer = v.type === "slicer" ? await v.getSlicerState().catch(() => null) : null;
            const vf = await v.getFilters().catch(() => []);
            return { name: v.name, type: v.type, title: v.title, x: Math.round(v.layout.x), y: Math.round(v.layout.y), width: Math.round(v.layout.width), height: Math.round(v.layout.height), hidden: v.layout.displayState?.mode === 1, ...(slicer ? { slicerOn: (slicer.targets ?? []).map(field) } : {}), ...(vf.length ? { filters: vf.map((f: any) => field(f.target)) } : {}) };
          }),
        ),
      });
    }
    const filters = await report.getFilters().catch(() => []);
    return { pages: out, reportFilters: filters.map((f: any) => ({ field: field(f.target), values: f.values, operator: f.operator })) };
  });
}

/* ---------------------------------- PDF ---------------------------------- */

/** Pictures of a PDF's pages (all, or those asked: 1-based), `width` px wide. */
export async function renderPdf(pdf: Uint8Array, pages?: number[], width = 1280): Promise<Buffer[]> {
  const page = await newPage();
  try {
    await page.setContent("<!doctype html><html><body></body></html>");
    const [lib, worker] = await Promise.all([script("pdfjs-dist/build/pdf.min.mjs"), script("pdfjs-dist/build/pdf.worker.min.mjs")]);
    const images: string[] = await page.evaluate(
      async ({ lib, worker, data, pages, width }) => {
        const url = (code: string) => URL.createObjectURL(new Blob([code], { type: "text/javascript" }));
        const pdfjs = await import(url(lib));
        pdfjs.GlobalWorkerOptions.workerSrc = url(worker);
        const doc = await pdfjs.getDocument({ data: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)) }).promise;
        const wanted = pages?.length ? pages : Array.from({ length: doc.numPages }, (_, i) => i + 1);
        const out: string[] = [];
        for (const n of wanted) {
          if (n < 1 || n > doc.numPages) continue;
          const p = await doc.getPage(n);
          const base = p.getViewport({ scale: 1 });
          const viewport = p.getViewport({ scale: width / base.width });
          const canvas = document.createElement("canvas");
          canvas.width = Math.round(viewport.width);
          canvas.height = Math.round(viewport.height);
          await p.render({ canvas, canvasContext: canvas.getContext("2d")!, viewport }).promise;
          out.push(canvas.toDataURL("image/png").split(",")[1]);
        }
        return out;
      },
      { lib, worker, data: Buffer.from(pdf).toString("base64"), pages, width },
    );
    return images.map((b) => Buffer.from(b, "base64"));
  } finally {
    await page.close().catch(() => undefined);
  }
}

/** A picture PowerPoint cannot take (SVG, WebP, AVIF…) drawn as a PNG (`width` px wide at most). */
export async function rasterize(bytes: Uint8Array, mime: string, width = 1200): Promise<Buffer> {
  const page = await newPage();
  try {
    await page.setContent("<!doctype html><html><body></body></html>");
    const png: string = await page.evaluate(
      async ({ data, mime, width }) => {
        const img = new Image();
        img.src = `data:${mime};base64,${data}`;
        await img.decode();
        const w0 = img.naturalWidth || 600, h0 = img.naturalHeight || 600;
        const k = Math.min(1, width / w0) * (w0 < 300 ? 300 / w0 : 1);
        const c = Object.assign(document.createElement("canvas"), { width: Math.round(w0 * k), height: Math.round(h0 * k) });
        c.getContext("2d")!.drawImage(img, 0, 0, c.width, c.height);
        return c.toDataURL("image/png").split(",")[1];
      },
      { data: Buffer.from(bytes).toString("base64"), mime, width },
    );
    return Buffer.from(png, "base64");
  } finally {
    await page.close().catch(() => undefined);
  }
}

/** Closes the headless browser (server stop). */
export async function closeBrowser() {
  const b = await browser?.catch(() => undefined);
  await b?.close().catch(() => undefined);
}
