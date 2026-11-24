/**
 * Runs chat turns from the command line (one process: the conversation is kept between turns):
 *   bun scripts/checks/agent.ts <screen> '<scope json>' '<message>' ['<answer>' …]
 */
import { runChat } from "@server/agent/run";
import { closeBrowser } from "@server/capture/browser";
import { settings } from "@server/settings";
// HUB_PROVIDER / HUB_MODEL: another provider for this check only (not saved).
const s = await settings();
if (process.env.HUB_PROVIDER) s.ai.provider = process.env.HUB_PROVIDER as never;
if (process.env.HUB_MODEL) s.ai.model = process.env.HUB_MODEL;
const [screen, scopeText, ...messages] = process.argv.slice(2);
const scope = JSON.parse(scopeText);
const conv = `check-${Date.now()}`;
for (const message of messages) {
  const t = Date.now();
  let text = "";
  console.log(`\n=== USER: ${message}`);
  await runChat({ conversationId: conv, screen: screen as never, message, scope }, (e) => {
    if (e.type === "text") text += e.text;
    else if (e.type === "tool" && e.status !== "running") console.log(`[${e.name} ${e.status} ${e.ms}ms${e.images ? ` 🖼${e.images}` : ""}] ${JSON.stringify(e.args).slice(0, 220)}\n   → ${(e.result ?? "").slice(0, 260).replace(/\n/g, "\n     ")}`);
    else if (e.type === "effect" && e.name === "ask") console.log(`[ASK] ${JSON.stringify(e.data).slice(0, 700)}`);
    else if (e.type === "effect" && e.name === "doc") ((scope.doc = e.data), console.log(`[doc updated] ${JSON.stringify(e.data).slice(0, 160)}`));
    else if (e.type === "effect") console.log(`[effect ${e.name}]`);
    else if (e.type === "client_tool") console.log(`[browser ${e.name}] (no browser here)`);
    else if (e.type === "error") console.log(`[ERROR] ${e.error}`);
  }, new AbortController().signal);
  console.log(`--- answer (${((Date.now() - t) / 1000).toFixed(0)} s) ---\n${text}`);
}
if (scope.doc) await Bun.write(".local/last-doc.json", JSON.stringify(scope.doc, null, 1));
await closeBrowser();
