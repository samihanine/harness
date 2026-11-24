/**
 * Claude through the Anthropic API (an API key): streamed, native tool calls, images, adaptive thinking,
 * prompt caching (the system prompt and tools are a stable prefix), server-side fallback on a refusal.
 */
import Anthropic from "@anthropic-ai/sdk";
import { AuthError, type ChatProvider, type CompleteRequest, type CompleteResult, type Message } from "./types";

/** Models that accept the server-side fallback on a refusal. */
const FALLBACK = /^claude-(fable-5|opus-5|sonnet-5-5)/;

function toParams(messages: Message[]): Anthropic.Beta.BetaMessageParam[] {
  return messages.map((m) => {
    if (m.role === "assistant" && m.raw?.provider === "anthropic") return { role: "assistant", content: m.raw.content as Anthropic.Beta.BetaContentBlockParam[] };
    const content: Anthropic.Beta.BetaContentBlockParam[] = m.content.map((b) => {
      switch (b.type) {
        case "text":
          return { type: "text", text: b.text || " " };
        case "image":
          return { type: "image", source: { type: "base64", media_type: b.mediaType as "image/png", data: b.data } };
        case "tool_call":
          return { type: "tool_use", id: b.id, name: b.name, input: b.args };
        case "tool_result":
          return {
            type: "tool_result",
            tool_use_id: b.id,
            is_error: b.isError || undefined,
            content: b.content.map((c) => (c.type === "text" ? { type: "text" as const, text: c.text || " " } : { type: "image" as const, source: { type: "base64" as const, media_type: c.mediaType as "image/png", data: c.data } })),
          };
      }
    });
    return { role: m.role, content };
  });
}

export function anthropicProvider(apiKey: string, baseURL?: string): ChatProvider {
  const client = new Anthropic({ apiKey, baseURL, maxRetries: 3 });
  return {
    id: "anthropic",
    async complete(req: CompleteRequest): Promise<CompleteResult> {
      const fallback = FALLBACK.test(req.model);
      const params: Anthropic.Beta.MessageCreateParamsStreaming = {
        model: req.model,
        max_tokens: 32000,
        stream: true,
        cache_control: { type: "ephemeral" },
        system: [{ type: "text", text: req.system }],
        tools: req.tools.map((t) => ({ name: t.name, description: t.description, input_schema: t.schema as Anthropic.Beta.BetaTool.InputSchema })),
        messages: toParams(req.messages),
        ...(/haiku-4/.test(req.model) ? {} : { thinking: { type: "adaptive" }, output_config: { effort: req.effort } }),
        ...(fallback ? { betas: ["server-side-fallback-2026-07-01"], fallbacks: "default" } : {}),
      } as Anthropic.Beta.MessageCreateParamsStreaming;
      let stream;
      try {
        stream = client.beta.messages.stream(params, { signal: req.signal });
        stream.on("text", (t) => req.onText?.(t));
        const final = await stream.finalMessage();
        return toResult(final);
      } catch (e) {
        if (e instanceof Anthropic.AuthenticationError) throw new AuthError("The Anthropic API key was refused (Settings › AI).");
        // An account or model without the fallback beta: asked again without it.
        if (fallback && e instanceof Anthropic.BadRequestError && /fallback/i.test(e.message)) {
          const { betas: _b, fallbacks: _f, ...plain } = params as any;
          stream = client.beta.messages.stream(plain, { signal: req.signal });
          stream.on("text", (t) => req.onText?.(t));
          return toResult(await stream.finalMessage());
        }
        throw e;
      }
    },
  };
}

function toResult(final: Anthropic.Beta.BetaMessage): CompleteResult {
  const content: Message["content"] = [];
  for (const b of final.content) {
    if (b.type === "text") content.push({ type: "text", text: b.text });
    else if (b.type === "tool_use") content.push({ type: "tool_call", id: b.id, name: b.name, args: (b.input ?? {}) as Record<string, unknown> });
  }
  // Sent back exactly as received (thinking blocks included); a fallback block is not a content to send back.
  const raw = final.content.filter((b) => (b.type as string) !== "fallback");
  const stop = final.stop_reason === "tool_use" ? "tool_use" : final.stop_reason === "max_tokens" ? "max_tokens" : final.stop_reason === "refusal" ? "refusal" : "end";
  return { message: { role: "assistant", content, raw: { provider: "anthropic", content: raw } }, stop, usage: { input: final.usage.input_tokens + (final.usage.cache_read_input_tokens ?? 0) + (final.usage.cache_creation_input_tokens ?? 0), output: final.usage.output_tokens } };
}
