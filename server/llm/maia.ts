/**
 * Maia: the group's TrustLayer v2 (POST <url>/v2/chat/generation/, headers client_id / client_secret),
 * written from its documentation. One format for every model it serves: messages of "parts" (text,
 * base64 images, function calls), roles user / model / tool, tools as "toolsDeclarations". Not streamed.
 */
import { AuthError, resultText, type ChatProvider, type CompleteRequest, type CompleteResult, type Message } from "./types";

export type MaiaConfig = { url: string; clientId: string; clientSecret: string; maison: string; useCaseId: string; environment: string };

/** The TrustLayer's provider of a model name. */
export function maiaProvider(model: string) {
  if (/^deployment-/.test(model)) return "AZURE";
  if (/^gpt-image|dall-e/.test(model)) return "DALLE";
  if (/^gpt-/.test(model)) return "OPENAI";
  if (/^gemini/.test(model)) return "GEMINI";
  if (/^imagen/.test(model)) return "IMAGEN";
  if (/^claude/.test(model)) return "ANTHROPIC";
  if (/^mistral/.test(model)) return "MISTRAL";
  return "AZURE";
}

/** Chat models of the TrustLayer (its documented list, image models left out). */
export const MAIA_MODELS = [
  "claude-opus-5-5", "claude-sonnet-5-5", "claude-sonnet-5", "claude-opus-4-8", "claude-sonnet-4-6",
  "gpt-6.1-sol", "gpt-6-astra", "gpt-6-sol", "gpt-6-luna", "gpt-5.6-sol", "gpt-5.6-terra", "gpt-5.6-luna",
  "deployment-gpt-5.4-2026-03-05", "deployment-gpt-5.2-2025-12-11", "deployment-gpt4.1-2025-04-14",
  "gemini-4-argon", "gemini-3.8-flash", "gemini-3.7-flash", "gemini-3.5-flash", "gemini-3.1-pro-preview", "gemini-2.5-pro", "mistral-medium-3",
];

function toMaia(messages: Message[]) {
  const out: { role: string; parts: unknown[] }[] = [];
  for (const m of messages) {
    if (m.role === "assistant") {
      const parts = m.content.flatMap((b): unknown[] => (b.type === "text" && b.text ? [{ type: "text", text: b.text }] : b.type === "tool_call" ? [{ type: "functionCall", id: b.id, name: b.name, parameters: b.args }] : []));
      if (parts.length) out.push({ role: "model", parts });
      continue;
    }
    const results = m.content.filter((b) => b.type === "tool_result");
    if (results.length) {
      // Function responses are objects: a text result goes as { result }.
      out.push({ role: "tool", parts: results.map((r) => ({ name: r.name, id: r.id, response: { [r.isError ? "error" : "result"]: resultText(r) } })) });
      // Pictures returned by tools follow as a user message (function responses take no image).
      const images = results.flatMap((r) => r.content.filter((c) => c.type === "image"));
      if (images.length) out.push({ role: "user", parts: [{ type: "text", text: "Pictures returned by the tools above:" }, ...images.map((i) => ({ data: (i as any).data, mimeType: (i as any).mediaType }))] });
    }
    const parts = m.content.flatMap((b): unknown[] => (b.type === "text" ? [{ type: "text", text: b.text }] : b.type === "image" ? [{ data: b.data, mimeType: b.mediaType }] : []));
    if (parts.length) out.push({ role: "user", parts });
  }
  return out;
}

/** The JSON schema as the TrustLayer's declarations take it (no $schema / additionalProperties). */
function clean(schema: unknown): unknown {
  if (Array.isArray(schema)) return schema.map(clean);
  if (!schema || typeof schema !== "object") return schema;
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(schema)) if (k !== "$schema" && k !== "additionalProperties") out[k] = clean(v);
  return out;
}

export function maiaChat(config: MaiaConfig, conversationId: () => string): ChatProvider {
  let step = 0;
  return {
    id: "maia",
    async complete(req: CompleteRequest): Promise<CompleteResult> {
      if (!config.clientId || !config.clientSecret) throw new AuthError("Maia needs a client id and a client secret (Settings › AI).");
      const body = {
        provider: maiaProvider(req.model),
        model: req.model,
        systemInstruction: req.system,
        messages: toMaia(req.messages),
        ...(req.tools.length ? { tools: [{ toolsDeclarations: req.tools.map((t) => ({ name: t.name, description: t.description, parameters: clean(t.schema) })) }] } : {}),
        params: { maxOutputTokens: 16000, stream: false },
        infos: { environment: config.environment, conversationId: conversationId(), maison: config.maison, useCaseId: config.useCaseId, callId: crypto.randomUUID(), callStep: String(step++) },
      };
      const r = await fetch(`${config.url.replace(/\/+$/, "")}/v2/chat/generation/`, {
        method: "POST",
        signal: req.signal,
        headers: { client_id: config.clientId, client_secret: config.clientSecret, accept: "application/json", "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const text = await r.text();
      if (r.status === 401 || r.status === 403) throw new AuthError(`Maia refused the credentials (HTTP ${r.status}).`);
      if (r.status === 451) throw new Error("Maia: the answer was blocked by the guardrails (451).");
      if (!r.ok) throw new Error(`Maia error ${r.status}: ${text.slice(0, 400)}`);
      const json = JSON.parse(text);
      const parts: any[] = json.message?.parts ?? [];
      const content: Message["content"] = [];
      for (const p of parts) {
        if (p.type === "functionCall" || p.functionCall) {
          const call = p.functionCall ?? p;
          content.push({ type: "tool_call", id: call.id ?? `call_${crypto.randomUUID().slice(0, 8)}`, name: call.name, args: call.parameters ?? call.args ?? {} });
        } else if (typeof p.text === "string") {
          content.push({ type: "text", text: p.text });
          req.onText?.(p.text);
        }
      }
      return {
        message: { role: "assistant", content },
        stop: content.some((c) => c.type === "tool_call") ? "tool_use" : "end",
        usage: json.usage ? { input: json.usage.inputTokens ?? 0, output: json.usage.outputTokens ?? 0 } : undefined,
      };
    },
  };
}
