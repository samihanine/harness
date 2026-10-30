/**
 * run_js: the AI's own JavaScript, run in a throwaway Web Worker with no network (fetch, XHR,
 * WebSocket, imports removed) and a time limit. It only gets the data it is given and returns a result:
 * it can compute, never write anywhere.
 */
const PRELUDE = `for (const k of ["fetch", "XMLHttpRequest", "WebSocket", "EventSource", "importScripts", "indexedDB", "caches", "BroadcastChannel"]) { try { Object.defineProperty(self, k, { value: undefined }); } catch {} }`;

export function runSandboxed(code: string, input: unknown, ms = 5000): Promise<{ result?: string; log: string[]; error?: string }> {
  const source = `${PRELUDE}
self.onmessage = async (event) => {
  const data = event.data;
  const log = [];
  const console = { log: (...a) => log.push(a.map((x) => (typeof x === "string" ? x : JSON.stringify(x))).join(" ")) };
  try {
    const result = await (async (data, console) => { ${code}\n })(data, console);
    postMessage({ result: result === undefined ? undefined : JSON.stringify(result), log });
  } catch (e) {
    postMessage({ error: String(e && e.stack ? e.message : e), log });
  }
};`;
  const url = URL.createObjectURL(new Blob([source], { type: "text/javascript" }));
  const worker = new Worker(url);
  return new Promise((resolve) => {
    const timer = setTimeout(() => {
      worker.terminate();
      resolve({ log: [], error: `Stopped after ${ms / 1000} s.` });
    }, ms);
    worker.onmessage = (e) => {
      clearTimeout(timer);
      worker.terminate();
      URL.revokeObjectURL(url);
      resolve(e.data);
    };
    worker.onerror = (e) => {
      clearTimeout(timer);
      worker.terminate();
      resolve({ log: [], error: e.message });
    };
    worker.postMessage(input);
  });
}
