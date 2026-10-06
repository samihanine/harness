/**
 * The app side of the agent: an MCP server (official SDK) connected to the agent iframe.
 * Pages add their own tools / resources / prompts while they are open (see useAgentFeatures).
 */
import { useEffect, useRef } from "react";
import { PostMessageTransport } from "@modelcontextprotocol/ext-apps";
import { McpServer } from "@modelcontextprotocol/server";

/** URL of the ia-agent app. */
export const AGENT_URL = import.meta.env.VITE_AGENT_URL ?? "http://localhost:3200/";

export const server = new McpServer(
  { name: "excel-agent", title: "Excel agent", version: "1.0.0" },
  {
    instructions: `The user edits local Excel files through schemas (field definitions).
The resource app://page says which page is open and what can be done there.
Values must match the field types; options must be one of the allowed values.`,
  },
);

// Handlers are created on first registration and must exist before connecting:
// register then remove one of each so pages can add theirs at any time.
for (const item of [
  server.registerTool("_init", { description: "" }, () => ({ content: [] })),
  server.registerResource("_init", "init://", {}, () => ({ contents: [] })),
  server.registerPrompt("_init", {}, () => ({ messages: [] })),
])
  item.remove();

const page = { text: "No page open." };
server.registerResource(
  "page",
  "app://page",
  { title: "Current page", mimeType: "text/plain", annotations: { priority: 1 } },
  async (uri) => ({ contents: [{ uri: uri.href, text: page.text }] }),
);

let connected: Window | null = null;
/** Connects the server to the agent iframe (once per iframe window). */
export async function connectAgent(target: Window) {
  if (connected === target) return;
  if (connected) await server.close();
  connected = target;
  await server.connect(new PostMessageTransport(target, target));
}

type Registered = { remove(): void };

/**
 * Adds features to the agent while the calling component is mounted.
 * `pageText` describes the page (resource app://page). Callbacks should read state through refs.
 */
export function useAgentFeatures(pageText: string, setup: (s: McpServer) => Registered[], deps: unknown[]) {
  useEffect(() => {
    page.text = pageText;
  }, [pageText]);
  useEffect(() => {
    const items = setup(server);
    return () => items.forEach((item) => item.remove());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps);
}

/** Ref always holding the latest value (for MCP callbacks registered once). */
export function useLatest<T>(value: T) {
  const ref = useRef(value);
  ref.current = value;
  return ref;
}

export const textResult = (value: unknown, undo?: { name: string; arguments: Record<string, unknown> }) => ({
  content: [{ type: "text" as const, text: typeof value === "string" ? value : JSON.stringify(value) }],
  ...(undo ? { _meta: { undo } } : {}),
});

export const errorResult = (error: unknown) => ({
  content: [{ type: "text" as const, text: error instanceof Error ? error.message : String(error) }],
  isError: true,
});
