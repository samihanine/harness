/**
 * The LLM layer: one message format for every provider. Anthropic and Maia (the TrustLayer) answer one
 * model call at a time (`complete`) and the agent loop (server/agent/run.ts) runs the tools between calls;
 * Claude Code runs its own loop with the same tools (server/llm/claude-code.ts).
 */

export type Image = { mediaType: string; data: string };
export type TextBlock = { type: "text"; text: string };
export type ImageBlock = { type: "image" } & Image;
export type ToolCall = { type: "tool_call"; id: string; name: string; args: Record<string, unknown> };
export type ToolResultBlock = { type: "tool_result"; id: string; name: string; content: (TextBlock | ImageBlock)[]; isError?: boolean };
export type Block = TextBlock | ImageBlock | ToolCall | ToolResultBlock;

/**
 * `raw`: the reply exactly as the provider sent it (e.g. Anthropic's thinking blocks), sent back
 * unchanged to the same provider: the conversation stays append-only.
 */
export type Message = { role: "user" | "assistant"; content: Block[]; raw?: { provider: string; content: unknown } };

export type ToolSpec = { name: string; description: string; schema: Record<string, unknown> };

export type CompleteRequest = {
  system: string;
  messages: Message[];
  tools: ToolSpec[];
  model: string;
  effort: "low" | "medium" | "high";
  signal: AbortSignal;
  /** Text as it is written (streaming providers). */
  onText?: (text: string) => void;
};

export type CompleteResult = { message: Message; stop: "end" | "tool_use" | "max_tokens" | "refusal"; usage?: { input: number; output: number } };

export interface ChatProvider {
  id: string;
  complete(req: CompleteRequest): Promise<CompleteResult>;
}

export class AuthError extends Error {}

/** A tool result's content as text (for providers that only take text there). */
export const resultText = (r: ToolResultBlock) => r.content.filter((c): c is TextBlock => c.type === "text").map((c) => c.text).join("\n");
