/**
 * The Hub server (Bun): serves the app and its API. The AI is called from here, with the screen's tools;
 * Microsoft (Power BI, Fabric, Graph) and Confluence are called from here with the user's sign-in and keys,
 * which never reach the browser (except the Power BI token the embeds need).
 *
 *   bun server/main.ts            API on PORT (default 3301; `bun dev` puts Vite in front on 3300)
 *   bun run build && bun server/main.ts --prod   the built app and the API on one port
 *
 * It listens on 127.0.0.1 and refuses requests from another site (CSRF, DNS rebinding). HOST and
 * HUB_ALLOWED_HOSTS open it to a network (no user accounts yet: one Microsoft sign-in for all).
 */
import type { ChatEvent, ChatRequest, SettingsPatch } from "@shared/types";
import { deckToPptx } from "./pptx/deck";
import { resolveZones } from "./pptx/pbi-zones";
import { templateGallery } from "./pptx/templates";
import { deckZones, parseDeck } from "@shared/deck";
import { slidePictures } from "./pptx/render";
import { authStatus, embedToken, signIn, signOut } from "./ms/auth";
import { pbi, scope } from "./ms/api";
import { browserResult, forgetConversation, runChat } from "./agent/run";
import { cleanFiles, fileInfo, readStoredFile, saveFile } from "./files";
import { folderIdOf, publicSettings, resetSettings, settings, updateSettings } from "./settings";
import { modelsOf, providerFor } from "./llm";
import { datasetCatalog, discoverDataset, discoverReport, reportCatalog } from "./pbi/catalog";
import { reportsAccess, reportAccess } from "./pbi/access";
import { reportStructure } from "./pbi/report-structure";
import { modelInfo } from "./pbi/models";
import { createBlankReport } from "./pbi/new-report";
import { listPages } from "./confluence";
import { analysisToExcel } from "./excel-export";
import { engineFor } from "./data/engine";
import { closeBrowser, warmReport } from "./capture/browser";
import type { Analysis } from "@shared/analysis";
import { hasApp, staticFile } from "./static";

/** Production: the built app is served too (always in the executable, which embeds it). */
const PROD = process.argv.includes("--prod") || process.env.NODE_ENV === "production" || (!process.argv[1]?.endsWith(".ts") && hasApp());
const PORT = Number(process.env.PORT ?? (PROD ? 3300 : 3301));
const HOST = process.env.HOST ?? "127.0.0.1";

/** Hosts the API answers (the page's own address); others are refused. */
const allowed = new Set([`127.0.0.1:${PORT}`, `localhost:${PORT}`, "127.0.0.1:3300", "localhost:3300", ...(process.env.HUB_ALLOWED_HOSTS ?? "").split(",").map((h) => h.trim()).filter(Boolean)]);
function sameOrigin(req: Request) {
  const host = req.headers.get("host");
  if (!host || !allowed.has(host)) return false;
  const origin = req.headers.get("origin");
  return !origin || allowed.has(origin.replace(/^https?:\/\//, ""));
}

const json = (body: unknown, status = 200) => Response.json(body, { status, headers: { "Cache-Control": "no-store" } });
const fail = (e: unknown, status = 500) => json({ error: e instanceof Error ? e.message : String(e) }, status);
const ref = (u: URL) => ({ id: u.searchParams.get("id") ?? "", groupId: u.searchParams.get("groupId") || undefined, appId: u.searchParams.get("appId") || undefined });
const GUID = /^[0-9a-f-]{36}$/i;

function chat(req: Request) {
  return req.json().then((body: ChatRequest) => {
    const abort = new AbortController();
    req.signal.addEventListener("abort", () => abort.abort());
    const enc = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {
        let closed = false;
        const emit = (e: ChatEvent) => {
          if (!closed) controller.enqueue(enc.encode(`data: ${JSON.stringify(e)}\n\n`));
        };
        // A comment every 15 s keeps proxies from closing a long silent turn.
        const ping = setInterval(() => !closed && controller.enqueue(enc.encode(": ping\n\n")), 15_000);
        try {
          await runChat(body, emit, abort.signal);
        } catch (e) {
          emit({ type: "error", error: e instanceof Error ? e.message : String(e) });
        } finally {
          closed = true;
          clearInterval(ping);
          try {
            controller.close();
          } catch {}
        }
      },
      cancel() {
        abort.abort();
      },
    });
    return new Response(stream, { headers: { "Content-Type": "text/event-stream", "Cache-Control": "no-cache", "X-Accel-Buffering": "no" } });
  });
}

async function api(req: Request, url: URL): Promise<Response> {
  const p = url.pathname;
  const m = req.method;
  // Auth (Microsoft)
  if (p === "/api/auth" && m === "GET") return json(await authStatus());
  if (p === "/api/auth/sign-in" && m === "POST") return json(await signIn());
  if (p === "/api/auth/sign-out" && m === "POST") return (await signOut(), json({ ok: true }));
  if (p === "/api/embed-token" && m === "GET") return json(await embedToken());
  // Settings
  if (p === "/api/settings" && m === "GET") return json({ ...(await publicSettings()), models: modelsOf((await settings()).ai.provider) });
  if (p === "/api/settings/reset" && m === "POST") {
    const s = await resetSettings(((await req.json()) as { part: "general" | "ai" | "powerbi" | "confluence" }).part);
    return json({ ...s, models: modelsOf((await settings()).ai.provider) });
  }
  if (p === "/api/settings/test-ai" && m === "POST") {
    const s = await settings();
    const t = Date.now();
    if (s.ai.provider === "claude-code") {
      const { runClaudeCode } = await import("./llm/claude-code");
      let text = "";
      await runClaudeCode({ conv: { system: "Answer with one word." }, tools: [], runTool: async () => ({ output: "", isError: false }), emit: (e) => e.type === "text" && (text += e.text), signal: AbortSignal.timeout(90_000), model: s.ai.model, effort: "low", maxSteps: 1, config: s.ai.claudeCode, user: { role: "user", content: [{ type: "text", text: "Say OK." }] } });
      return json({ ok: true, text: `Claude Code answered "${text.trim().slice(0, 40)}" with ${s.ai.model} (${Date.now() - t} ms).` });
    }
    const r = await providerFor(s, () => "connection-test").complete({ system: "Answer with one word.", messages: [{ role: "user", content: [{ type: "text", text: "Say OK." }] }], tools: [], model: s.ai.model, effort: "low", signal: AbortSignal.timeout(60_000) });
    const text = r.message.content.filter((b) => b.type === "text").map((b) => (b as { text: string }).text).join("");
    return json({ ok: true, text: `${s.ai.provider} answered "${text.trim().slice(0, 40)}" with ${s.ai.model} (${Date.now() - t} ms).` });
  }
  if (p === "/api/settings" && m === "PUT") {
    const s = await updateSettings((await req.json()) as SettingsPatch);
    return json({ ...s, models: modelsOf((await settings()).ai.provider) });
  }
  // Chat
  if (p === "/api/chat" && m === "POST") return chat(req);
  if (p === "/api/chat/tool-result" && m === "POST") {
    const b = (await req.json()) as { callId: string; result?: unknown; error?: string };
    return json({ ok: browserResult(b.callId, b.result, b.error) });
  }
  if (p.startsWith("/api/chat/") && m === "DELETE") return json({ ok: forgetConversation(p.split("/")[3]) });
  // Files
  if (p === "/api/files" && m === "POST") {
    const form = await req.formData();
    const file = form.get("file");
    if (!(file instanceof Blob)) return fail("No file", 400);
    if (file.size > 100 * 1024 * 1024) return fail("File larger than 100 MB", 413);
    const id = String(form.get("id") ?? "") || undefined;
    return json(await saveFile(new Uint8Array(await file.arrayBuffer()), (file as File).name ?? "file", file.type, id));
  }
  if (p.startsWith("/api/files/")) {
    const id = p.split("/")[3];
    if (m === "HEAD") return new Response(null, { status: (await fileInfo(id)) ? 200 : 404 });
    const { ref: f, data } = await readStoredFile(id);
    const inline = f.type.startsWith("image/") || url.searchParams.has("inline");
    return new Response(data as unknown as BodyInit, { headers: { "Content-Type": f.type, "Content-Disposition": `${inline ? "inline" : "attachment"}; filename*=UTF-8''${encodeURIComponent(f.name)}`, "Cache-Control": "private, max-age=86400" } });
  }
  // Power BI
  if (p === "/api/powerbi/reports") return json(await reportCatalog());
  if (p === "/api/powerbi/datasets") return json(await datasetCatalog());
  if (p === "/api/powerbi/report") {
    const r = ref(url);
    if (!GUID.test(r.id)) return fail("Bad report id", 400);
    const meta = r.appId ? await pbi(`/apps/${r.appId}/reports/${r.id}`) : await pbi(`${scope(r.groupId)}/reports/${r.id}`);
    const ds = meta.datasetId ? await pbi(`/datasets/${meta.datasetId}`).catch(() => undefined) : undefined;
    if (r.appId) discoverReport({ id: r.id, name: meta.name, appId: r.appId, workspace: "App", datasetId: meta.datasetId, webUrl: meta.webUrl });
    return json({ id: r.id, groupId: r.groupId, appId: r.appId, name: meta.name, datasetId: meta.datasetId, datasetGroupId: ds?.workspaceId ?? undefined, webUrl: meta.webUrl, access: await reportAccess(r) });
  }
  if (p === "/api/powerbi/dataset") {
    // A model from a pasted link (also an app's): its name, and it joins the list of models.
    const r = ref(url);
    if (!GUID.test(r.id)) return fail("Bad model id", 400);
    const meta = await pbi(`${scope(r.groupId)}/datasets/${r.id}`).catch(() => pbi(`/datasets/${r.id}`));
    discoverDataset({ id: r.id, name: meta.name, groupId: r.groupId, workspace: r.groupId ? "Workspace" : "Shared" });
    return json({ id: r.id, groupId: r.groupId, name: meta.name });
  }
  if (p === "/api/powerbi/access" && m === "POST") return json(await reportsAccess((await req.json()) as { id: string; groupId?: string; appId?: string }[]));
  if (p === "/api/powerbi/test") {
    const t = Date.now();
    const [reports, models] = await Promise.all([reportCatalog(), datasetCatalog()]);
    return json({ ok: true, text: `${reports.length} reports and ${models.length} semantic models reachable (${Date.now() - t} ms).` });
  }
  if (p === "/api/powerbi/warm" && m === "POST") {
    // The report on screen opened in the headless browser ahead of time (the first picture is then quick).
    const b = (await req.json()) as { id: string; groupId?: string; appId?: string };
    if (GUID.test(b.id)) warmReport(b);
    return json({ ok: true });
  }
  if (p === "/api/powerbi/report-structure") return json(await reportStructure(ref(url), url.searchParams.get("detail") === "full" ? "full" : "summary", true));
  if (p === "/api/powerbi/model") {
    const r = ref(url);
    return json(await modelInfo(r, url.searchParams.has("fresh")));
  }
  if (p === "/api/dax" && m === "POST") {
    const b = (await req.json()) as { id: string; groupId?: string; query: string };
    const { dax } = await import("./pbi/dax");
    return json([...(await dax({ id: b.id, groupId: b.groupId }, b.query))]);
  }
  if (p === "/api/powerbi/new-report" && m === "POST") {
    const b = (await req.json()) as { name: string; datasetId: string; groupId?: string };
    return json(await createBlankReport(b.name, { id: b.datasetId, groupId: b.groupId }, b.groupId));
  }
  // Confluence (projects)
  if (p === "/api/confluence/pages" && m === "GET") return json(await listPages(url.searchParams.get("folder") || undefined));
  if (p === "/api/projects/test" && m === "POST") {
    const { folderId } = (await req.json()) as { folderId: string };
    const pages = await listPages(folderIdOf(folderId));
    return json({ ok: true, text: `${pages.length} pages: ${pages.slice(0, 5).map((x) => x.title).join(", ")}${pages.length > 5 ? "…" : ""}` });
  }
  // Presenter
  if (p === "/api/deck/pptx" && m === "POST") {
    const b = (await req.json()) as { markdown: string; name?: string };
    const pptx = await deckToPptx(b.markdown);
    return new Response(pptx as unknown as BodyInit, { headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation", "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(b.name ?? "deck")}.pptx` } });
  }
  if (p === "/api/deck/preview" && m === "POST") {
    const b = (await req.json()) as { markdown?: string; fileId?: string; slides?: number[] };
    const data = b.fileId ? (await readStoredFile(b.fileId)).data : new Uint8Array(await deckToPptx(b.markdown ?? ""));
    const pics = await slidePictures(data, b.slides, 1280);
    const files = await Promise.all(pics.map((png, i) => saveFile(png, `slide-${i + 1}.png`, "image/png")));
    return json(files.map((f) => f.id));
  }
  if (p === "/api/deck/zones" && m === "POST") {
    // The Power BI views of a deck, resolved for live embeds (presentation mode).
    const { markdown } = (await req.json()) as { markdown: string };
    return json(await resolveZones(deckZones(parseDeck(markdown))));
  }
  if (p === "/api/deck/templates" && m === "GET") return json(await templateGallery());
  if (p === "/api/pptx/structure" && m === "GET") {
    const { Pptx } = await import("./pptx/ooxml");
    return json(await (await Pptx.load((await readStoredFile(url.searchParams.get("fileId") ?? "")).data)).structure());
  }
  if (p === "/api/pptx/add-template" && m === "POST") {
    // A slide drawn from the template added to an imported PowerPoint: a new version of the file.
    const b = (await req.json()) as { fileId: string; name: string; markdown: string; after: number };
    const { Pptx } = await import("./pptx/ooxml");
    const target = await Pptx.load((await readStoredFile(b.fileId)).data);
    const md = /^\s*layout\s*:/m.test(b.markdown.split("---")[0] ?? "") ? b.markdown : `layout: default\n---\n${b.markdown}`;
    const drawn = await Pptx.load(new Uint8Array(await deckToPptx(`---\ntitle: ${b.name}\n---\n${md}`)));
    await target.importSlide(drawn, 1, b.after);
    return json(await saveFile(await target.save(), b.name, "application/vnd.openxmlformats-officedocument.presentationml.presentation"));
  }
  // Data Scientist
  if (p === "/api/analysis/export" && m === "POST") {
    const a = (await req.json()) as Analysis;
    return new Response((await analysisToExcel(a)) as unknown as BodyInit, { headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(a.title)}.xlsx` } });
  }
  if (p === "/api/analysis/refresh" && m === "POST") {
    // Sources read again (fresh data), every block computed again.
    const a = (await req.json()) as Analysis;
    const scratch: Record<string, any> = {};
    await engineFor(scratch, a);
    const { select } = await import("./data/engine");
    const blocks = a.blocks.map((b) => {
      if (b.type === "text") return b;
      try {
        const r = select(scratch.engine, b.sql, 5000);
        return { ...b, rows: r.rows, ...(b.type === "table" ? { columns: r.columns } : {}) };
      } catch {
        return b;
      }
    });
    return json({ ...a, blocks, updatedAt: new Date().toISOString() });
  }
  return fail(`No route ${m} ${p}`, 404);
}

async function serveStatic(url: URL) {
  const path = decodeURIComponent(url.pathname).replace(/\.\.+/g, "").replace(/^\/+/, "");
  const file = path ? staticFile(`dist/${path}`) : undefined;
  if (file) return new Response(file, path.startsWith("assets/") ? { headers: { "Cache-Control": "public, max-age=31536000, immutable" } } : undefined);
  return new Response(staticFile("dist/index.html"), { headers: { "Content-Type": "text/html" } });
}

const server = Bun.serve({
  hostname: HOST,
  port: PORT,
  idleTimeout: 255,
  async fetch(req) {
    const url = new URL(req.url);
    if (url.pathname.startsWith("/api/")) {
      if (!sameOrigin(req)) return new Response("forbidden", { status: 403 });
      try {
        return await api(req, url);
      } catch (e) {
        const status = (e as { status?: number }).status;
        return fail(e, e instanceof Error && /Not signed in/.test(e.message) ? 401 : status && status >= 400 && status < 600 ? status : 500);
      }
    }
    if (PROD && hasApp()) return serveStatic(url);
    return new Response("Hub API. The app: bun dev → http://localhost:3300", { status: 404 });
  },
});

void cleanFiles();
setInterval(() => void cleanFiles(), 86_400_000).unref?.();
for (const sig of ["SIGINT", "SIGTERM"] as const)
  process.on(sig, async () => {
    await closeBrowser();
    process.exit(0);
  });
console.log(`Hub server on http://${HOST}:${server.port}${PROD ? "" : " (API; the app runs on :3300 with bun dev)"}`);
