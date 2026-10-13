/**
 * The agent loop for a plain-text LLM: every reply is one JSON object listing tool calls
 * (`{"tools":[{"name","args"}]}`), parsed tolerantly; the application runs them and sends the results.
 * One LLM conversation per app conversation: the protocol, tools and context are sent once, then each
 * user message only brings what changed (the page state, the context when it changed).
 */
import { z } from "zod";
import { openSession, type Session } from "./llm";
import { ReplyError, parseReply } from "./reply";

export type Tool = {
  name: string;
  description: string;
  args: z.ZodObject;
  /** Reads only (no confirmation needed, safe to repeat). */
  readOnly?: boolean;
  run: (args: any) => Promise<string>;
};

export const tool = <S extends z.ZodObject>(t: { name: string; description: string; args: S; readOnly?: boolean; run: (args: z.infer<S>) => Promise<string> }): Tool => t;

export type Agent = {
  /** What this screen is and how to work in it. */
  instructions: string;
  tools: Tool[];
  /** What the AI works on (model, file, notes…): sent at the start of a conversation, again only when it changes. */
  context: () => Promise<string>;
  /** What the user sees now (page, filters, editor…): re-read before each model call, sent when it changed. */
  state?: () => Promise<string>;
};

export type Step = { id: string; name: string; args: Record<string, unknown>; result?: string; error?: boolean; ms?: number };
/** One exchange with the model, exactly as sent and received. */
export type LlmCall = { at: string; prompt: string; reply?: string; error?: string; ms?: number };
export type Update = { steps: Step[]; llm: LlmCall[]; answer?: string; /** Remote conversation id (stateful APIs). */ remote?: string };

/** Readable error, also for the plain objects the Power BI client rejects with ({ message, detailedMessage }). */
export function describeError(error: unknown): string {
  if (error instanceof Error) return error.message;
  if (error && typeof error === "object") {
    const e = error as Record<string, unknown>;
    const text = [e.message, e.detailedMessage, (e.body as Record<string, unknown>)?.message, (e.body as Record<string, unknown>)?.detailedMessage].filter((x) => typeof x === "string" && x);
    return text.length ? [...new Set(text)].join(": ") : JSON.stringify(error).slice(0, 500);
  }
  return String(error);
}

/** sessionChars: past this size the LLM conversation is restarted (with a summary of the app conversation). */
const LIMITS = { steps: 20, retries: 2, resultChars: 6_000, historyChars: 12_000, sessionChars: 150_000 };

const tag = (name: string, body: string, attrs = "") => `<${name}${attrs ? ` ${attrs}` : ""}>\n${body.trim()}\n</${name}>`;

const PROTOCOL = `You are the engine of an AI assistant embedded in an application. Messages are written by the APPLICATION: it gives you the context and the user's request, runs the tools you call and sends you their results. The user only sees what you pass to "answer".

YOU HAVE TOOLS. Calling them is the ONLY way to read data or change anything in the application: you never act by describing an action in text. To use a tool, reply with its name and arguments in JSON (format below); the application runs it and sends you the result in its next message.

1. Every reply is ONE raw JSON object and nothing else (no prose, no markdown fences): {"tools": [{"name": "<tool>", "args": {…}}]}
2. Several independent tools may be called in one reply; you get all results in the next message.
3. When done, reply with "answer" ALONE: {"tools": [{"name": "answer", "args": {"text": "…"}}]}
4. Never invent data: read it with tools. Never claim a change you did not make through a tool.
5. When a tool fails, read the error, fix the arguments and retry, or explain the problem.
6. Do the whole request in one go: when it is clear, do every step with tools then answer; ask the user only when the request is truly ambiguous. When unsure about a fact, check it with a tool rather than guess.
7. Answer in the user's language, concise, in markdown.`;

/** End of EVERY message: weak models forget that they can (and must) call tools. */
const format = (names: string[]) =>
  tag(
    "reply_format",
    `IMPORTANT — reply with ONE raw JSON object and NOTHING else (no text before or after it, no markdown fences):
{"tools": [{"name": "<tool name>", "args": {…}}]}
Tool names: ${names.join(", ")}.
Need data, or must change something? Call the tool(s) NOW, e.g. {"tools": [{"name": "${names.find((n) => n !== "answer") ?? "answer"}", "args": {…}}]}
Done (everything was done through tools)? {"tools": [{"name": "answer", "args": {"text": "your reply to the user, in markdown"}}]}`,
  );

/** Live LLM conversation of each app conversation, with what it was last told (context, state). */
type Live = { session: Session; model: string; context?: string; state?: string };
const live = new Map<string, Live>();
/** Drops the LLM conversation of an app conversation (deleted). */
export const forgetSession = (conversationId: string) => void live.delete(conversationId);

/** `name(a: string, b?: number)` + descriptions, from the zod schema. */
function signature(t: Tool) {
  const schema = z.toJSONSchema(t.args) as { properties?: Record<string, any>; required?: string[] };
  const required = new Set(schema.required ?? []);
  const props = Object.entries(schema.properties ?? {});
  const type = (s: any): string =>
    s.enum ? s.enum.map((v: unknown) => JSON.stringify(v)).join(" | ") : s.anyOf ? s.anyOf.map(type).join(" | ") : s.type === "array" ? `${type(s.items ?? {})}[]` : s.type === "object" && s.properties ? `{ ${Object.entries(s.properties).map(([k, v]) => `${k}: ${type(v)}`).join("; ")} }` : (s.type ?? "any");
  return [
    `- ${t.name}(${props.map(([k, s]) => `${k}${required.has(k) ? "" : "?"}: ${type(s)}`).join(", ")})${t.readOnly ? "" : "  [changes things]"}`,
    `    ${t.description}`,
    ...props.filter(([, s]) => s.description).map(([k, s]) => `    ${k}: ${s.description}`),
  ].join("\n");
}

const ANSWER: Tool = tool({ name: "answer", description: "Final reply to the user (markdown). Ends the turn. Must be the only call of its reply.", args: z.object({ text: z.string() }), readOnly: true, run: async () => "" });

export async function* runAgent({
  agent,
  model,
  conversationId,
  history,
  request,
  signal,
}: {
  agent: Agent;
  model: string;
  /** App conversation: its LLM conversation is reused from one user message to the next. */
  conversationId: string;
  /** The app conversation so far, as text: only used to start a new LLM conversation (e.g. after a reload). */
  history: string;
  request: string;
  signal: AbortSignal;
}): AsyncGenerator<Update> {
  const tools = [ANSWER, ...agent.tools];
  const byName = new Map(tools.map((t) => [t.name, t]));
  const FORMAT = format(tools.map((t) => t.name));
  const steps: Step[] = [];
  const llm: LlmCall[] = [];
  const read = async () => ({
    context: await agent.context().catch((e) => `Context unavailable: ${e}`),
    state: agent.state ? await agent.state().catch((e) => `Page state unavailable: ${e}`) : undefined,
  });
  let current = live.get(conversationId);
  if (current && (current.model !== model || current.session.chars > LIMITS.sessionChars)) current = undefined;
  const fresh = !current;
  if (!current) live.set(conversationId, (current = { session: openSession(model, "agent"), model }));
  const conv = current;
  const now = await read();
  /** What changed since the LLM was last told, as tags (both on a new conversation). */
  const changes = (next: typeof now, first = false) => {
    const parts = [
      (first || next.context !== conv.context) && tag(first ? "context" : "context_now", next.context),
      next.state !== undefined && (first || next.state !== conv.state) && tag(first ? "page_state" : "page_state_now", next.state),
    ];
    conv.context = next.context;
    conv.state = next.state;
    return parts;
  };
  const unchanged = !fresh && now.state !== undefined && now.state === conv.state;
  let prompt = [
    ...(fresh ? [tag("protocol", PROTOCOL), tag("app_instructions", agent.instructions), tag("tools", tools.map(signature).join("\n"))] : []),
    ...changes(now, fresh),
    fresh && history && tag("conversation_so_far", history.slice(-LIMITS.historyChars)),
    unchanged && tag("note", "The page state has not changed since your last message."),
    tag("user_message", request),
    FORMAT,
  ]
    .filter(Boolean)
    .join("\n\n");
  const session = conv.session;
  let remote: string | undefined;
  let retries = 0;
  let reminded = false;
  try {
    for (let turn = 0; turn <= LIMITS.steps; turn++) {
      signal.throwIfAborted();
      const call: LlmCall = { at: new Date().toISOString(), prompt: turn === LIMITS.steps ? `${tag("note", "No tool turns left: answer now.")}\n\n${FORMAT}` : prompt };
      llm.push(call);
      const started = performance.now();
      const reply = await session.send(call.prompt, signal).catch((error: unknown) => {
        call.error = error instanceof Error ? error.message : String(error);
        throw error;
      });
      Object.assign(call, { reply, ms: Math.round(performance.now() - started) });
      remote ??= await session.remote();
      let calls;
      try {
        calls = parseReply(reply);
        const unknown = calls.filter((c) => !byName.has(c.name)).map((c) => c.name);
        if (unknown.length) throw new ReplyError(`Unknown tool(s): ${unknown.join(", ")}. Available: ${[...byName.keys()].join(", ")}.`);
        retries = 0;
      } catch (error) {
        // Plain prose (no JSON at all) is the model answering: taken as the final answer, no extra round trip.
        if (error instanceof ReplyError && !reply.includes("{") && reply.trim()) return yield { steps, llm, remote, answer: reply.trim() };
        if (!(error instanceof ReplyError) || ++retries > LIMITS.retries) return yield { steps, llm, remote, answer: reply.trim() };
        prompt = `${tag("invalid_reply", `Your reply could not be used: ${error.message}\nSend the corrected JSON only.`)}\n\n${FORMAT}`;
        turn--;
        continue;
      }
      const answer = calls.find((c) => c.name === "answer");
      if (answer && calls.length === 1) {
        // Failed steps never redone successfully: one reminder before the answer, so the model retries
        // them or says what failed instead of claiming them done.
        const failed = steps.filter((s, i) => s.error && !steps.slice(i + 1).some((t) => t.name === s.name && !t.error));
        if (failed.length && !reminded) {
          reminded = true;
          prompt = [
            tag(
              "note",
              `Before answering: these steps FAILED and were not redone successfully:\n${failed.map((s) => `- ${s.name} ${JSON.stringify(s.args).slice(0, 200)} → ${String(s.result).slice(0, 300)}`).join("\n")}\nFix the arguments and retry them now, or say clearly in your answer what is not done. Never present them as done.`,
            ),
            FORMAT,
          ].join("\n\n");
          continue;
        }
        return yield { steps, llm, remote, answer: String(answer.args.text ?? "") };
      }

      const results: string[] = [];
      for (const call of calls.filter((c) => c.name !== "answer")) {
        const step: Step = { id: crypto.randomUUID(), name: call.name, args: call.args };
        steps.push(step);
        yield { steps, llm, remote };
        const t = byName.get(call.name)!;
        const started = performance.now();
        try {
          const text = await t.run(t.args.parse(call.args));
          step.result = text.length > LIMITS.resultChars ? `${text.slice(0, LIMITS.resultChars)}\n…[truncated, ${text.length} characters: ask for less]` : text;
        } catch (error) {
          if (signal.aborted) throw error;
          step.error = true;
          step.result = error instanceof z.ZodError ? `Invalid arguments: ${z.prettifyError(error)}` : describeError(error);
        }
        step.ms = Math.round(performance.now() - started);
        results.push(tag("result", step.result || "(done)", `tool="${call.name}" status="${step.error ? "error" : "ok"}"`));
        yield { steps, llm, remote };
      }
      prompt = [tag("tool_results", results.join("\n")), answer && tag("note", `"answer" was ignored: it must be alone in its reply.`), ...changes(await read()), FORMAT]
        .filter(Boolean)
        .join("\n\n");
    }
  } catch (error) {
    // A failed exchange may have left the LLM conversation incomplete: the next message starts a new one.
    if (live.get(conversationId) === conv) live.delete(conversationId);
    throw error;
  }
  // Remote conversations are kept: deleted with the app conversation (see Chat).
}
