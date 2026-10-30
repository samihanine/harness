/**
 * Screenshot of a screen of the app with given library entries, in a headless Chrome with its own
 * profile (.local/agent-eval/chrome): the dev server (bun dev, port 3300) gives it the Power BI
 * embed tokens, so reports render as the user would see them.
 *
 *   bun scripts/agent-eval/render.ts <seed.json> <path, e.g. /viewer?report=…> <out.png> [wait seconds] [script.js]
 *   script.js (optional): an async expression run in the page once loaded (its result is printed),
 *   e.g. to drive the AI's tools against the embedded report; the screenshot is taken after it.
 *   seed.json: { "reports": [...], "datasets": [...], ... } (IndexedDB collections of the app)
 */
import { spawn } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";

const [seedFile, path, out, wait = "18", script] = process.argv.slice(2);
if (!out) throw new Error("bun scripts/agent-eval/render.ts <seed.json> <path> <out.png> [wait seconds]");
const seed = JSON.parse(readFileSync(seedFile, "utf8"));
const APP = "http://localhost:3300";
const PORT = 9334;

const chrome = spawn("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", ["--headless=new", `--remote-debugging-port=${PORT}`, `--user-data-dir=${process.cwd()}/.local/agent-eval/chrome`, "--window-size=1600,1000", "--hide-scrollbars", "about:blank"], { stdio: "ignore" });
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
try {
  let target: { webSocketDebuggerUrl: string } | undefined;
  for (let i = 0; i < 40 && !target; i++) {
    await sleep(250);
    target = await fetch(`http://127.0.0.1:${PORT}/json`).then((r) => r.json()).then((l: any[]) => l.find((t) => t.type === "page")).catch(() => undefined);
  }
  const ws = new WebSocket(target!.webSocketDebuggerUrl);
  await new Promise((r) => (ws.onopen = r));
  let id = 0;
  const waiting = new Map<number, (v: any) => void>();
  const errors: string[] = [];
  ws.onmessage = (m) => {
    const d = JSON.parse(String(m.data));
    if (d.id && waiting.has(d.id)) (waiting.get(d.id)!(d), waiting.delete(d.id));
    if (d.method === "Runtime.exceptionThrown") errors.push(String(d.params.exceptionDetails.exception?.description ?? d.params.exceptionDetails.text).slice(0, 300));
  };
  const send = (method: string, params = {}) => new Promise<any>((r) => { const i = ++id; waiting.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
  await send("Runtime.enable");
  await send("Page.enable");
  await send("Emulation.setDeviceMetricsOverride", { width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false });
  // The app's origin first (IndexedDB is per origin), then the collections, then the screen.
  await send("Page.navigate", { url: `${APP}/settings` });
  await sleep(4000);
  const put = `(async () => { const db = await new Promise((r, j) => { const q = indexedDB.open("keyval-store"); q.onupgradeneeded = () => q.result.createObjectStore("keyval"); q.onsuccess = () => r(q.result); q.onerror = j; }); const seed = ${JSON.stringify(seed)}; await new Promise((r) => { const t = db.transaction("keyval", "readwrite"); for (const [k, v] of Object.entries(seed)) t.objectStore("keyval").put(v, k); t.oncomplete = r; }); return "ok"; })()`;
  await send("Runtime.evaluate", { expression: put, awaitPromise: true });
  await send("Page.navigate", { url: `${APP}${path}` });
  await sleep(Number(wait) * 1000);
  if (script) {
    const result = await send("Runtime.evaluate", { expression: readFileSync(script, "utf8"), awaitPromise: true, returnByValue: true });
    console.log(result.result?.exceptionDetails ? `Script error: ${JSON.stringify(result.result.exceptionDetails).slice(0, 800)}` : `Script result:\n${typeof result.result?.result?.value === "string" ? result.result.result.value : JSON.stringify(result.result?.result?.value, null, 2)}`);
    await sleep(6000);
  }
  const shot = await send("Page.captureScreenshot", { format: "png" });
  writeFileSync(out, Buffer.from(shot.result.data, "base64"));
  console.log(errors.length ? `Page errors:\n${errors.join("\n")}` : `Screenshot: ${out}`);
  ws.close();
} finally {
  chrome.kill();
}
