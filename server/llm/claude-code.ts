/**
 * Claude Code as the model (Claude Agent SDK): an isolated instance — no built-in tool (Bash, Read…), no
 * settings, hooks, plugins or CLAUDE.md read from the machine — that runs its own loop over the screen's
 * tools, given as an in-process MCP server. The connection is explicit: the gateway address, its headers
 * and the key helper (the Okta script of Claude Maison), or an API key; Zscaler certificates when present.
 * The conversation is resumed from its session id.
 */
import { createSdkMcpServer, query, tool as sdkTool } from "@anthropic-ai/claude-agent-sdk";
import { existsSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import type { ChatEvent } from "@shared/types";
import type { Tool, ToolOutput } from "@/tools/tool";
import { DATA_DIR } from "../settings";
import type { Message } from "./types";

type Config = { baseUrl: string; headers: string; keyHelper: string; key: string };

const CERT_DIR = join(homedir(), ".zscaler-certificates");
const certificate = (envValue: string | undefined, file: string) => envValue ?? (existsSync(join(CERT_DIR, file)) ? join(CERT_DIR, file).replaceAll("\\", "/") : undefined);

/** The sub-process environment: nothing ANTHROPIC_* / CLAUDE* inherited, the strict necessary given. */
function claudeEnv(c: Config, model: string) {
  const env: Record<string, string | undefined> = {};
  for (const [k, v] of Object.entries(process.env)) if (!/^(ANTHROPIC_|CLAUDE)/i.test(k)) env[k] = v;
  const rootCa = certificate(process.env.NODE_EXTRA_CA_CERTS, "zscaler-root-ca.pem");
  const bundle = certificate(process.env.SSL_CERT_FILE, "ca-bundle-with-zscaler.pem");
  return {
    ...env,
    ...(c.baseUrl ? { ANTHROPIC_BASE_URL: c.baseUrl } : {}),
    ...(c.headers ? { ANTHROPIC_CUSTOM_HEADERS: c.headers } : {}),
    ...(c.key ? { ANTHROPIC_API_KEY: c.key } : !c.keyHelper && (process.env.ANTHROPIC_API_KEY || process.env.CLAUDE_API_KEY) ? { ANTHROPIC_API_KEY: process.env.ANTHROPIC_API_KEY ?? process.env.CLAUDE_API_KEY } : {}),
    ANTHROPIC_MODEL: model,
    ...(rootCa ? { NODE_EXTRA_CA_CERTS: rootCa } : {}),
    ...(bundle ? { SSL_CERT_FILE: bundle } : {}),
    CLAUDE_CODE_DISABLE_EXPERIMENTAL_BETAS: "1",
    CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC: "1",
    CLAUDE_CONFIG_DIR: join(DATA_DIR, "claude-code"),
  };
}

/** The key helper as Claude Code runs it: the script through bash (Git for Windows' bash on Windows). */
function keyHelper(script: string) {
  if (!script) return undefined;
  const bash =
    Bun.which("bash") ?? [join(homedir(), "AppData", "Local", "Programs", "Git", "bin", "bash.exe"), "C:\\Program Files\\Git\\bin\\bash.exe"].find((p) => existsSync(p)) ?? "bash";
  return `"${bash.replaceAll("\\", "/")}" "${script.replaceAll("\\", "/")}"`;
}

const toMcp = (o: ToolOutput, isError: boolean) => ({
  isError,
  content: typeof o === "string" ? [{ type: "text" as const, text: o || "(done)" }] : [{ type: "text" as const, text: o.text || "(done)" }, ...(o.images ?? []).map((i) => ({ type: "image" as const, data: i.data, mimeType: i.mediaType }))],
});

export async function runClaudeCode(o: {
  conv: { system: string; claudeSession?: string };
  tools: Tool[];
  runTool: (name: string, args: Record<string, unknown>, id: string) => Promise<{ output: ToolOutput; isError: boolean }>;
  emit: (e: ChatEvent) => void;
  signal: AbortSignal;
  model: string;
  effort: "low" | "medium" | "high";
  maxSteps: number;
  config: Config;
  user: Message;
}) {
  const server = createSdkMcpServer({
    name: "hub",
    version: "2.0.0",
    tools: o.tools.map((t): any =>
      sdkTool(t.name, t.description, t.args.shape as never, async (args: Record<string, unknown>) => {
        const r = await o.runTool(t.name, args, crypto.randomUUID());
        return toMcp(r.output, r.isError);
      }),
    ),
  });
  const names = o.tools.map((t) => `mcp__hub__${t.name}`);
  const abort = new AbortController();
  o.signal.addEventListener("abort", () => abort.abort());
  const content = o.user.content.map((b) => (b.type === "image" ? { type: "image", source: { type: "base64", media_type: b.mediaType, data: b.data } } : { type: "text", text: (b as { text: string }).text }));
  async function* prompt() {
    yield { type: "user" as const, message: { role: "user" as const, content }, parent_tool_use_id: null, session_id: o.conv.claudeSession ?? "" };
  }
  const helper = keyHelper(o.config.keyHelper);
  const q = query({
    prompt: prompt() as never,
    options: {
      cwd: join(DATA_DIR, "claude-code", "workspace"),
      abortController: abort,
      model: o.model,
      effort: o.effort,
      ...(o.conv.claudeSession ? { resume: o.conv.claudeSession } : {}),
      env: claudeEnv(o.config, o.model),
      settingSources: [],
      ...(helper ? { settings: { apiKeyHelper: helper } } : {}),
      systemPrompt: o.conv.system,
      tools: [],
      mcpServers: { hub: server },
      strictMcpConfig: true,
      // Only the screen's tools: anything else is refused.
      canUseTool: async (name: string, input: Record<string, unknown>) => (names.includes(name) ? { behavior: "allow" as const, updatedInput: input } : { behavior: "deny" as const, message: `Tool not allowed: ${name}` }),
      includePartialMessages: true,
      maxTurns: o.maxSteps,
    } as never,
  });
  let session = o.conv.claudeSession;
  for await (const msg of q as AsyncIterable<any>) {
    if (msg.type === "system" && msg.subtype === "init") session = msg.session_id;
    else if (msg.type === "stream_event" && msg.event?.type === "content_block_delta" && msg.event.delta?.type === "text_delta") o.emit({ type: "text", text: msg.event.delta.text });
    else if (msg.type === "result" && msg.subtype !== "success") throw new Error((msg.errors ?? [msg.subtype]).join("; "));
  }
  return { session };
}
