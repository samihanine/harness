/** Runs one chat turn from the command line: bun scripts/checks/agent.ts <screen> '<scope json>' '<message>' [conversationId] */
import { runChat } from "@server/agent/run";
import { closeBrowser } from "@server/capture/browser";
const [screen, scope, message, conv = "check-" + Date.now()] = process.argv.slice(2);
const t = Date.now();
let text = "";
await runChat({ conversationId: conv, screen: screen as never, message, scope: JSON.parse(scope) }, (e) => {
  if (e.type === "text") text += e.text;
  else if (e.type === "tool" && e.status !== "running") console.log(`[tool ${e.name} ${e.status} ${e.ms}ms] ${JSON.stringify(e.args).slice(0, 200)}\n   → ${(e.result ?? "").slice(0, 300).replace(/\n/g, "\n     ")}`);
  else if (e.type === "effect") console.log(`[effect ${e.name}] ${JSON.stringify(e.data).slice(0, 300)}`);
  else if (e.type === "client_tool") console.log(`[browser ${e.name}] (no browser in this check)`);
  else if (e.type === "error") console.log(`[error] ${e.error}`);
}, new AbortController().signal);
console.log(`\n--- answer (${Date.now() - t} ms) ---\n${text}`);
await closeBrowser();
