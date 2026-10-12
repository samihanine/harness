/**
 * AI calls with the server's key (AI_KEY in .env), used when no key is set in the app's settings:
 * the key never reaches the browser. Headers come from the browser with placeholders
 * ({{AI_KEY}}, {{AI_KEY:NAME}}) replaced here, so llm.ts stays the only place describing the provider.
 */
import { createServerFn } from "@tanstack/react-start";
import { AI_BASE_URL } from "@/agent/ai-base";

const parse = (raw: string) => {
  const fields = Object.fromEntries([...raw.matchAll(/([A-Za-z0-9_]+)=(\S+)/g)].map((m) => [m[1], m[2]]));
  return { token: Object.values(fields)[0] ?? raw.trim(), fields };
};

export const aiProxy = createServerFn({ method: "POST" })
  .inputValidator((d: { path: string; method: string; headers: Record<string, string>; body?: unknown }) => d)
  .handler(async ({ data }) => {
    const raw = process.env.AI_KEY;
    if (!raw) return { status: 0, body: "", type: "" };
    const key = parse(raw);
    const headers = Object.fromEntries(
      Object.entries(data.headers).map(([k, v]) => [k, v.replace(/\{\{AI_KEY(?::(\w+))?\}\}/g, (_, field) => (field ? (key.fields[field] ?? "") : key.token))]),
    );
    const r = await fetch(`${AI_BASE_URL}${data.path}`, {
      method: data.method,
      headers: { "Content-Type": "application/json", ...headers },
      body: data.body === undefined ? undefined : JSON.stringify(data.body),
    });
    return { status: r.status, body: await r.text(), type: r.headers.get("content-type") ?? "" };
  });
