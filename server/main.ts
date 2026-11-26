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
import { authStatus, embedToken, signIn, signOut } from "./ms/auth";
import { pbi, scope } from "./ms/api";
import { browserResult, forgetConversation, runChat } from "./agent/run";
import { cleanFiles, fileInfo, readStoredFile, saveFile } from "./files";
import { folderIdOf, publicSettings, resetSettings, settings, updateSettings } from "./settings";
import { modelsOf, providerFor } from "./llm";
import { projectsList } from "./projects";
import { saveGif, startTeams, teamsChats, teamsStatus, updateTeams, type TeamsConfig } from "./teams/bot";
import { deckPdf, deckPptx, powerbiPicture } from "./slides/render";
import { parseLink } from "@shared/links";
import type { Template } from "@shared/slides";
import { addFileSheets, closeSession, createWorkbook, ensureConnections, workbookBytes } from "./workbook/excel";
import { sheetHtml, sheetPreview } from "./workbook/preview";
import { forgetPbixCopy, importPbix, pbixInfo, pbixWithDefinition } from "./pbi/pbix";
import { resolveDataset, resolveReport } from "./pbi/resolve";
import { addToLibrary, changeLibrary, library, removeFromLibrary } from "./library";
import { datasetCatalog, discoverDataset, discoverReport, reportCatalog } from "./pbi/catalog";
import { reportStructure } from "./pbi/report-structure";
import { modelInfo } from "./pbi/models";
import { folderChildren, listPages } from "./confluence";
import { closeBrowser, warmReport } from "./capture/browser";
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
    return json(await resolveReport(r));
  }
  if (p === "/api/powerbi/dataset") {
    // A model from a link (also an app's): its name and workspace.
    const r = ref(url);
    if (!GUID.test(r.id)) return fail("Bad model id", 400);
    return json(await resolveDataset(r));
  }
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

  // Data Scientist: the analysis's workbook in OneDrive (Excel computes it), shown as HTML, downloaded as is.
  if (p === "/api/workbook/create" && m === "POST") {
    const b = (await req.json()) as { key: string; title: string };
    return json(await createWorkbook(b.key, b.title));
  }
  if (p === "/api/workbook/preview" && m === "GET") {
    const preview = await sheetPreview(url.searchParams.get("itemId") ?? "", url.searchParams.get("sheet") ?? undefined);
    return json({ sheets: preview.sheets, sheet: preview.sheet, total: preview.total, shown: { rows: preview.cells.length, cols: preview.widths.length }, html: sheetHtml(preview) });
  }
  if (p === "/api/workbook/connect" && m === "POST") {
    const b = (await req.json()) as { itemId: string; models: { id: string; name: string }[] };
    return json({ connections: await ensureConnections(b.itemId, b.models) });
  }
  if (p === "/api/workbook/add-file" && m === "POST") {
    const b = (await req.json()) as { itemId: string; fileId: string };
    const f = await readStoredFile(b.fileId);
    return json({ sheets: await addFileSheets(b.itemId, { name: f.ref.name, data: f.data }) });
  }
  if (p === "/api/workbook/download" && m === "GET") {
    const id = url.searchParams.get("itemId") ?? "";
    await closeSession(id);
    return new Response((await workbookBytes(id)) as unknown as BodyInit, { headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } });
  }
  // Teams: the assistant of one chat (settings, its activity, the user's chats, the mascot's GIFs made by the app).
  if (p === "/api/teams" && m === "GET") return json(await teamsStatus());
  if (p === "/api/teams" && m === "POST") return json(await updateTeams((await req.json()) as Partial<TeamsConfig>));
  if (p === "/api/teams/chats" && m === "GET") return json(await teamsChats());
  if (p.startsWith("/api/teams/gif/") && m === "POST") return json((await saveGif(p.split("/").pop()!, new Uint8Array(await req.arrayBuffer())), { ok: true }));
  // Builder: a dropped .pbix imported privately (My workspace) to be shown and changed; downloaded with its changes.
  if (p === "/api/pbix/open" && m === "POST") {
    const b = (await req.json()) as { fileId: string; name: string; key: string };
    const data = (await readStoredFile(b.fileId)).data;
    const info = await pbixInfo(data);
    const copy = await importPbix(data, `[Hub] ${b.name.replace(/\.pbix$/i, "")} · ${b.key.slice(0, 6)}`);
    return json({ ...info, ...copy });
  }
  if (p === "/api/pbix/save" && m === "POST") {
    // The dropped file with the copy's current definition: downloaded, and kept as the item's latest version.
    const b = (await req.json()) as { fileId: string; reportId: string; name: string };
    const out = await pbixWithDefinition((await readStoredFile(b.fileId)).data, { id: b.reportId });
    return json(await saveFile(Buffer.from(out), b.name.endsWith(".pbix") ? b.name : `${b.name}.pbix`, "application/octet-stream"));
  }
  if (p === "/api/pbix/forget" && m === "POST") return json((await forgetPbixCopy(((await req.json()) as { reportId: string }).reportId), { ok: true }));
  // Confluence: projects (sub-folders of the projects folder) and the catalog of reports and models (a page each).
  if (p === "/api/projects" && m === "GET") return json(await projectsList(url.searchParams.get("refresh") ?? undefined));
  if (p === "/api/confluence/test" && m === "POST") {
    const { folderId } = (await req.json()) as { folderId: string };
    const children = await folderChildren(folderIdOf(folderId));
    return json({ ok: true, text: `${children.length} items: ${children.slice(0, 6).map((x) => x.title).join(", ")}${children.length > 6 ? "…" : ""}` });
  }
  const lib = p.match(/^\/api\/library\/(report|dataset)(?:\/(\d+))?$/);
  if (lib) {
    const kind = lib[1] as "report" | "dataset";
    if (m === "GET") return json(await library(kind, url.searchParams.get("fresh") === "1"));
    if (m === "POST" && !lib[2]) return json(await addToLibrary(kind, (await req.json()) as { link: string; description?: string }));
    if (m === "PATCH" && lib[2]) return json(await changeLibrary(kind, lib[2], (await req.json()) as { name?: string; description?: string }));
    if (m === "DELETE" && lib[2]) return json(await removeFromLibrary(kind, lib[2]));
  }
  // Presenter: templates × markdown, drawn as the app draws them (pictures, PDF) or as a native PowerPoint.
  if (p.startsWith("/api/assets/template/") && m === "GET") {
    const f = staticFile(`assets/template/${p.split("/").pop()!.replace(/[^\w.-]/g, "")}`);
    return f ? new Response(f, { headers: { "Cache-Control": "max-age=86400" } }) : fail("No such picture", 404);
  }
  if (p === "/api/slides/powerbi" && m === "GET") {
    const png = await powerbiPicture(url.searchParams.get("url") ?? "");
    return new Response(png as unknown as BodyInit, { headers: { "Content-Type": "image/png", "Cache-Control": "max-age=600" } });
  }
  if (p === "/api/slides/pptx" && m === "POST") {
    const b = (await req.json()) as { markdown: string; templates: Template[]; title?: string };
    return new Response((await deckPptx(b.markdown, b.templates, b.title)) as unknown as BodyInit, { headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation" } });
  }
  if (p === "/api/slides/pdf" && m === "POST") {
    const b = (await req.json()) as { markdown: string; templates: Template[] };
    return new Response((await deckPdf(b.markdown, b.templates)) as unknown as BodyInit, { headers: { "Content-Type": "application/pdf" } });
  }
  if (p === "/api/slides/resolve" && m === "POST") {
    // Power BI links of slides → the report's ids, page and visual names, for live views in the presentation.
    const { urls } = (await req.json()) as { urls: string[] };
    return json(await Promise.all(urls.map(async (u) => {
      const l = parseLink(u);
      const id = l?.report ?? l?.bare;
      if (!id) return { url: u, error: "Not a report link" };
      const r = await resolveReport({ id, groupId: l?.groupId, appId: l?.appId }).catch((e) => ({ error: e instanceof Error ? e.message : String(e) }) as never);
      return { url: u, ref: { id: r.id, groupId: r.groupId, appId: r.appId }, pageName: l?.page, visualName: l?.visual };
    })));
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
startTeams();
setInterval(() => void cleanFiles(), 86_400_000).unref?.();
for (const sig of ["SIGINT", "SIGTERM"] as const)
  process.on(sig, async () => {
    await closeBrowser();
    process.exit(0);
  });
console.log(`Hub server on http://${HOST}:${server.port}${PROD ? "" : " (API; the app runs on :3300 with bun dev)"}`);
