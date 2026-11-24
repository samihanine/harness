/**
 * The agent of every screen. One conversation per chat, kept on the server between messages (append-only,
 * so the provider's cache and thinking carry over). Each message: the model is called with the screen's
 * tools, the tools run, the results go back, until it answers or asks the user (`ask_user`: questions with
 * buttons, which ends the turn). The first message of a conversation always gets a clarification question
 * before anything is changed. The state the AI works on is never pasted into the prompt: it reads it with
 * the tools (report state, screenshots, structures…).
 */
import { z } from "zod";
import type { ChatEvent, ChatRequest } from "@shared/types";
import type { ConversationState, Tool, ToolContext, ToolOutput } from "@/tools/tool";
import { askUser } from "@/tools/common/ask-user";
import { readStoredFile } from "../files";
import { settings } from "../settings";
import { providerFor } from "../llm";
import { AuthError, type Block, type Message, type ToolSpec } from "../llm/types";
import { runClaudeCode } from "../llm/claude-code";
import { screenAgent } from "./screens";
import { FIRST_QUESTION_REMINDER, NOT_ASKED_YET, scopeText } from "@/prompts/common";

type Conversation = { state: ConversationState; messages: Message[]; system: string; scopeKey: string; claudeSession?: string; provider: string; tokens: number; lastAt: number };
const conversations = new Map<string, Conversation>();

/** Past this many input tokens a conversation starts again from its transcript (calls stay fast). */
const RESTART_TOKENS = 350_000;
const RESULT_CHARS = 40_000;

/** Browser actions waiting for their result (POST /api/chat/tool-result). */
const waiting = new Map<string, { resolve: (v: unknown) => void; reject: (e: Error) => void }>();
export function browserResult(callId: string, result: unknown, error?: string) {
  const w = waiting.get(callId);
  if (!w) return false;
  waiting.delete(callId);
  error ? w.reject(new Error(error)) : w.resolve(result);
  return true;
}

/** Forgets conversations idle for hours (their transcript comes back from the browser if needed). */
setInterval(() => {
  for (const [id, c] of conversations) if (Date.now() - c.lastAt > 6 * 3600_000) conversations.delete(id);
}, 600_000).unref?.();

export const forgetConversation = (id: string) => conversations.delete(id);

const specOf = (t: Tool): ToolSpec => {
  const schema = z.toJSONSchema(t.args) as Record<string, unknown>;
  delete schema.$schema;
  return { name: t.name, description: t.description, schema };
};

const asText = (o: ToolOutput) => (typeof o === "string" ? o : o.text);
const cut = (t: string) => (t.length > RESULT_CHARS ? `${t.slice(0, RESULT_CHARS)}\n…[cut at ${RESULT_CHARS} characters]` : t);

/** Readable message of anything thrown (the Power BI client rejects plain objects). */
export function describeError(e: unknown): string {
  if (e instanceof Error) return e.message;
  if (e && typeof e === "object") {
    const o = e as Record<string, any>;
    return [o.message, o.detailedMessage].filter(Boolean).join(": ") || JSON.stringify(e).slice(0, 400);
  }
  return String(e);
}

export async function runChat(req: ChatRequest, emit: (e: ChatEvent) => void, signal: AbortSignal) {
  const s = await settings();
  const agent = screenAgent(req.screen);
  const tools = [askUser, ...agent.tools(req.scope)];
  const scopeKey = scopeText(req.scope);
  let conv = conversations.get(req.conversationId);
  const provider = s.ai.provider;
  if (conv && conv.provider !== provider) conv = undefined; // another provider: restarted from the transcript
  if (!conv) {
    conv = {
      state: { id: req.conversationId, screen: req.screen, asked: !!req.transcript, scratch: {} },
      messages: [],
      system: agent.system(req.scope),
      scopeKey: "",
      provider,
      tokens: 0,
      lastAt: Date.now(),
    };
    conversations.set(req.conversationId, conv);
  }
  conv.lastAt = Date.now();
  // A question was asked at the end of the last turn: this message is its answer.
  if (conv.state.scratch.askedPending) conv.state.asked = true;
  const runId = crypto.randomUUID();
  emit({ type: "start", runId, model: s.ai.model });

  const ctx: ToolContext = {
    scope: req.scope,
    conversation: conv.state,
    signal,
    effect: (name, data) => emit({ type: "effect", name, data }),
    status: (label) => emit({ type: "status", label }),
    browser: (action, args = {}) =>
      new Promise((resolve, reject) => {
        const callId = crypto.randomUUID();
        const timer = setTimeout(() => browserResult(callId, undefined, "The browser did not answer in 60 s (is the screen still open?)."), 60_000);
        waiting.set(callId, { resolve: (v) => (clearTimeout(timer), resolve(v)), reject: (e) => (clearTimeout(timer), reject(e)) });
        emit({ type: "client_tool", callId, name: action, args });
      }),
  };

  // The user's message: what changed on the screen, the message, its pictures and files.
  const content: Block[] = [];
  const notes: string[] = [];
  if (!conv.messages.length && req.transcript) notes.push(`Earlier in this conversation (the server restarted it):\n${req.transcript.slice(-20_000)}`);
  if (scopeKey !== conv.scopeKey) {
    notes.push(conv.scopeKey ? `The screen changed. Now:\n${scopeKey}` : `On the screen:\n${scopeKey}`);
    conv.scopeKey = scopeKey;
    const brief = await agent.briefing?.(ctx).catch(() => undefined);
    if (brief) notes.push(brief);
  }
  for (const f of req.files ?? []) {
    if (f.type.startsWith("image/")) {
      const { data } = await readStoredFile(f.id);
      content.push({ type: "image", mediaType: f.type, data: Buffer.from(data).toString("base64") });
      notes.push(`Picture attached: ${f.name}`);
    } else notes.push(`File attached: "${f.name}" (${f.type || "file"}, id ${f.id})`);
  }
  if (!conv.state.asked) notes.push(NOT_ASKED_YET);
  content.unshift({ type: "text", text: [...notes.map((n) => `<note>\n${n}\n</note>`), req.message].join("\n\n") });
  conv.messages.push({ role: "user", content });

  /** Runs one tool call: the first-message rule, arguments checked, output cut, every step shown. */
  const runTool = async (name: string, rawArgs: Record<string, unknown>, id: string): Promise<{ output: ToolOutput; isError: boolean }> => {
    const t = tools.find((x) => x.name === name);
    const started = Date.now();
    emit({ type: "tool", id, name, args: rawArgs, status: "running" });
    let output: ToolOutput;
    let isError = false;
    try {
      if (!t) throw new Error(`No tool "${name}". Tools: ${tools.map((x) => x.name).join(", ")}.`);
      if (t.changes && !conv!.state.asked) throw new Error(NOT_ASKED_YET);
      const parsed = t.args.safeParse(rawArgs ?? {});
      if (!parsed.success) throw new Error(`Wrong arguments: ${parsed.error.issues.map((i) => `${i.path.join(".") || "args"}: ${i.message}`).join("; ")}`);
      output = await t.run(parsed.data, ctx);
    } catch (e) {
      output = describeError(e);
      isError = true;
    }
    const text = cut(asText(output));
    const images = typeof output === "string" ? [] : (output.images ?? []);
    emit({ type: "tool", id, name, args: rawArgs, status: isError ? "error" : "done", result: text.slice(0, 4000), images: images.length, ms: Date.now() - started });
    return { output: typeof output === "string" ? text : { text, images }, isError };
  };

  try {
    if (provider === "claude-code") {
      const r = await runClaudeCode({ conv, tools, runTool, emit, signal, model: s.ai.model, effort: s.ai.effort, maxSteps: s.ai.maxSteps, config: s.ai.claudeCode, user: conv.messages.at(-1)! });
      conv.claudeSession = r.session;
      return;
    }
    const llm = providerFor(s, () => req.conversationId);
    const specs = tools.map(specOf);
    let reminded = false;
    for (let step = 0; ; step++) {
      if (signal.aborted) throw new Error("Stopped");
      if (step >= s.ai.maxSteps) {
        emit({ type: "text", text: `\n\n(Stopped after ${s.ai.maxSteps} steps: Settings › AI › maximum steps.)` });
        break;
      }
      const result = await llm.complete({ system: conv.system, messages: conv.messages, tools: specs, model: s.ai.model, effort: s.ai.effort, signal, onText: (t) => emit({ type: "text", text: t }) });
      conv.messages.push(result.message);
      conv.tokens = result.usage?.input ?? conv.tokens;
      if (result.stop === "refusal") {
        emit({ type: "text", text: "\n\n(The model declined this request.)" });
        break;
      }
      const calls = result.message.content.filter((b) => b.type === "tool_call");
      if (!calls.length) {
        // The first message of a conversation gets a clarification question: sent back once if missing.
        if (!conv.state.asked && !reminded) {
          reminded = true;
          conv.messages.push({ role: "user", content: [{ type: "text", text: FIRST_QUESTION_REMINDER }] });
          emit({ type: "text", text: "\n\n" });
          continue;
        }
        break;
      }
      // In the order given: reads that follow each other run side by side, a change waits for what came before it.
      const outputs = new Map<string, Awaited<ReturnType<typeof runTool>>>();
      let group: typeof calls = [];
      const flush = async () => {
        await Promise.all(group.map(async (c) => outputs.set(c.id, await runTool(c.name, c.args, c.id))));
        group = [];
      };
      for (const c of calls) {
        if (tools.find((t) => t.name === c.name)?.changes) {
          await flush();
          outputs.set(c.id, await runTool(c.name, c.args, c.id));
        } else group.push(c);
      }
      await flush();
      conv.messages.push({
        role: "user",
        content: calls.map((c) => {
          const { output, isError } = outputs.get(c.id)!;
          const blocks: Extract<Block, { type: "tool_result" }>["content"] = [{ type: "text", text: asText(output) || "(done)" }];
          if (typeof output !== "string") for (const i of output.images ?? []) blocks.push({ type: "image", mediaType: i.mediaType, data: i.data });
          return { type: "tool_result", id: c.id, name: c.name, content: blocks, isError };
        }),
      });
      // A question to the user ends the turn: the answer is the next message.
      if (calls.some((c) => c.name === askUser.name && !outputs.get(c.id)!.isError)) break;
    }
    if (conv.tokens > RESTART_TOKENS) restart(conv);
  } catch (e) {
    if (e instanceof AuthError) emit({ type: "error", error: e.message });
    else emit({ type: "error", error: signal.aborted ? "Stopped." : describeError(e) });
    // A failed call leaves the conversation open on a tool call: closed so the next message can follow.
    const last = conv.messages.at(-1);
    if (last?.role === "assistant" && last.content.some((b) => b.type === "tool_call")) conv.messages.pop();
    if (conv.messages.at(-1)?.role === "user" && conv.messages.length) conv.messages.push({ role: "assistant", content: [{ type: "text", text: "(interrupted)" }] });
    return;
  }
  emit({ type: "done", usage: conv.tokens ? { input: conv.tokens, output: 0 } : undefined });
}

/** A long conversation starts again with its text only (the drafts and data stay in the state). */
function restart(conv: Conversation) {
  const text = conv.messages
    .flatMap((m) => m.content.filter((b) => b.type === "text").map((b) => `${m.role === "user" ? "User" : "Assistant"}: ${(b as { text: string }).text}`))
    .join("\n\n")
    .slice(-30_000);
  conv.messages = [{ role: "user", content: [{ type: "text", text: `<note>\nThe conversation so far (restarted to stay fast):\n${text}\n</note>\nContinue from here.` }] }, { role: "assistant", content: [{ type: "text", text: "Understood." }] }];
  conv.tokens = 0;
}
