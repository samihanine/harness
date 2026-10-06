/**
 * Connection to the app hosting the agent, through the Model Context Protocol.
 *
 * The host page runs an MCP server (tools, resources, prompts, instructions) and connects it to
 * this iframe with the standard `PostMessageTransport`; the agent is the MCP client.
 *
 * Conventions on top of plain MCP:
 * - resources with `annotations.priority ≥ 0.8` are "inline": sent with every request and
 *   resent when they change. Others are listed and read on demand.
 * - a tool result may carry `_meta.undo = { name, arguments }`: the host tool call reverting it.
 * - tool `annotations.readOnlyHint` marks tools that change nothing (no approval, run in parallel).
 */
import { Client } from "@modelcontextprotocol/client";
import { PostMessageTransport } from "@modelcontextprotocol/ext-apps";
import type { UndoAction } from "@/lib/types";

export type JsonSchema = {
  type?: string | string[];
  description?: string;
  properties?: Record<string, JsonSchema>;
  required?: string[];
  items?: JsonSchema;
  enum?: unknown[];
  anyOf?: JsonSchema[];
  [key: string]: unknown;
};

export type HostTool = {
  name: string;
  description: string;
  schema: JsonSchema;
  readOnly: boolean;
};

export type HostResource = {
  uri: string;
  name: string;
  description?: string;
  mimeType?: string;
  inline: boolean;
  /** URI template (RFC 6570), e.g. `table://{name}/rows`. */
  template?: boolean;
};

export type HostPrompt = { name: string; title: string; description?: string };

export type ToolOutcome = { text: string; isError: boolean; undo?: UndoAction };

const CALL_TIMEOUT = 10 * 60_000;
const opts = { cacheMode: "bypass" as const, timeout: CALL_TIMEOUT };

export class Host {
  private constructor(private client: Client) {}

  private static connection?: Promise<Host | null>;
  private static listeners = new Set<() => void>();

  /** Connects (once) to the parent page; null when the agent is not embedded or no server answers. */
  static connect(onChange: () => void): Promise<Host | null> {
    Host.listeners.add(onChange);
    Host.connection ??= Host.open(() => Host.listeners.forEach((listener) => listener()));
    return Host.connection;
  }

  private static async open(onChange: () => void): Promise<Host | null> {
    if (window.parent === window) return null;
    const client = new Client(
      { name: "agent", version: "1.0.0" },
      {
        listChanged: {
          tools: { onChanged: onChange },
          resources: { onChanged: onChange },
          prompts: { onChanged: onChange },
        },
      },
    );
    const transport = new PostMessageTransport(window.parent, window.parent);
    const connected = client.connect(transport).then(() => true);
    const timeout = new Promise<boolean>((resolve) => setTimeout(() => resolve(false), 5_000));
    return (await Promise.race([connected, timeout])) ? new Host(client) : null;
  }

  get info() {
    const version = this.client.getServerVersion();
    return {
      name: version?.name ?? "host",
      title: version?.title ?? version?.name ?? "App",
      instructions: this.client.getInstructions() ?? "",
    };
  }

  private get capabilities() {
    return this.client.getServerCapabilities() ?? {};
  }

  async tools(): Promise<HostTool[]> {
    if (!this.capabilities.tools) return [];
    const { tools } = await this.client.listTools(undefined, opts);
    return tools.map((tool) => ({
      name: tool.name,
      description: tool.description ?? tool.title ?? "",
      schema: tool.inputSchema as JsonSchema,
      readOnly: !!tool.annotations?.readOnlyHint,
    }));
  }

  async resources(): Promise<HostResource[]> {
    if (!this.capabilities.resources) return [];
    const [{ resources }, { resourceTemplates }] = await Promise.all([
      this.client.listResources(undefined, opts),
      this.client.listResourceTemplates(undefined, opts).catch(() => ({ resourceTemplates: [] })),
    ]);
    return [
      ...resources.map((r) => ({
        uri: r.uri,
        name: r.title ?? r.name,
        description: r.description,
        mimeType: r.mimeType,
        inline: (r.annotations?.priority ?? 0) >= 0.8,
      })),
      ...resourceTemplates.map((t) => ({
        uri: t.uriTemplate,
        name: t.title ?? t.name,
        description: t.description,
        mimeType: t.mimeType,
        inline: false,
        template: true,
      })),
    ];
  }

  async prompts(): Promise<HostPrompt[]> {
    if (!this.capabilities.prompts) return [];
    const { prompts } = await this.client.listPrompts(undefined, opts);
    return prompts.map((p) => ({ name: p.name, title: p.title ?? p.name, description: p.description }));
  }

  async prompt(name: string) {
    const { messages } = await this.client.getPrompt({ name });
    return messages.map((m) => (m.content.type === "text" ? m.content.text : "")).join("\n");
  }

  async read(uri: string) {
    const { contents } = await this.client.readResource({ uri }, opts);
    return contents.map((c) => ("text" in c ? c.text : `[binary ${c.mimeType ?? ""}]`)).join("\n");
  }

  async call(name: string, args: Record<string, unknown>): Promise<ToolOutcome> {
    const result = await this.client.callTool({ name, arguments: args }, { timeout: CALL_TIMEOUT });
    const content = (result.content ?? []) as { type: string; text?: string }[];
    const text = content.map((c) => (c.type === "text" ? c.text : `[${c.type}]`)).join("\n");
    const undo = (result._meta as { undo?: UndoAction } | undefined)?.undo;
    return { text: text || JSON.stringify(result.structuredContent ?? {}), isError: !!result.isError, undo };
  }
}
