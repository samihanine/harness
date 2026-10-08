/**
 * The agent loop for a plain-text LLM: every reply is one JSON object listing tool calls
 * (`{"tools":[{"name","args"}]}`), parsed tolerantly; the application runs them and sends the results.
 * The page context (what the user sees) is sent with the request and again whenever it changes.
 */
import { z } from "zod";
import { openSession } from "./llm";
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
  /** What the user sees now (re-read before each model call). */
  context: () => Promise<string>;
};

export type Step = { id: string; name: string; args: Record<string, unknown>; result?: string; error?: boolean; ms?: number };
/** One exchange with the model, exactly as sent and received. */
export type LlmCall = { at: string; prompt: string; reply?: string; error?: string; ms?: number };
export type Update = { steps: Step[]; llm: LlmCall[]; answer?: string };

const LIMITS = { steps: 20, retries: 2, resultChars: 6_000, historyChars: 12_000 };

const tag = (name: string, body: string, attrs = "") => `<${name}${attrs ? ` ${attrs}` : ""}>\n${body.trim()}\n</${name}>`;

const PROTOCOL = `You are the engine of an AI assistant embedded in an application. Messages are written by the APPLICATION: it gives you the context and the user's request, runs the tools you call and sends you their results. The user only sees what you pass to "answer".

1. Every reply is ONE raw JSON object and nothing else (no prose, no markdown fences): {"tools": [{"name": "<tool>", "args": {…}}]}
2. Several independent tools may be called in one reply; you get all results in the next message.
3. When done, reply with "answer" ALONE: {"tools": [{"name": "answer", "args": {"text": "…"}}]}
4. Never invent data: read it with tools. Never claim a change you did not make through a tool.
5. When a tool fails, read the error, fix the arguments and retry, or explain the problem.
6. Answer in the user's language, concise, in markdown.`;

const FORMAT = tag("reply_format", `Reply now with ONE raw JSON object: {"tools": [{"name": …, "args": {…}}]}. Final reply: {"tools": [{"name": "answer", "args": {"text": "…"}}]}`);

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
  history,
  request,
  signal,
}: {
  agent: Agent;
  model: string;
  history: string;
  request: string;
  signal: AbortSignal;
}): AsyncGenerator<Update> {
  const tools = [ANSWER, ...agent.tools];
  const byName = new Map(tools.map((t) => [t.name, t]));
  const steps: Step[] = [];
  const llm: LlmCall[] = [];
  let context = await agent.context().catch((e) => `Context unavailable: ${e}`);
  let prompt = [
    tag("protocol", PROTOCOL),
    tag("app_instructions", agent.instructions),
    tag("tools", tools.map(signature).join("\n")),
    tag("context", context),
    history && tag("conversation_so_far", history.slice(-LIMITS.historyChars)),
    tag("user_message", request),
    FORMAT,
  ]
    .filter(Boolean)
    .join("\n\n");
  const session = openSession(model, "agent");
  let retries = 0;
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
      let calls;
      try {
        calls = parseReply(reply);
        const unknown = calls.filter((c) => !byName.has(c.name)).map((c) => c.name);
        if (unknown.length) throw new ReplyError(`Unknown tool(s): ${unknown.join(", ")}. Available: ${[...byName.keys()].join(", ")}.`);
        retries = 0;
      } catch (error) {
        if (!(error instanceof ReplyError) || ++retries > LIMITS.retries) return yield { steps, llm, answer: reply.trim() };
        prompt = `${tag("invalid_reply", `Your reply could not be used: ${error.message}\nSend the corrected JSON only.`)}\n\n${FORMAT}`;
        turn--;
        continue;
      }
      const answer = calls.find((c) => c.name === "answer");
      if (answer && calls.length === 1) return yield { steps, llm, answer: String(answer.args.text ?? "") };

      const results: string[] = [];
      for (const call of calls.filter((c) => c.name !== "answer")) {
        const step: Step = { id: crypto.randomUUID(), name: call.name, args: call.args };
        steps.push(step);
        yield { steps, llm };
        const t = byName.get(call.name)!;
        const started = performance.now();
        try {
          const text = await t.run(t.args.parse(call.args));
          step.result = text.length > LIMITS.resultChars ? `${text.slice(0, LIMITS.resultChars)}\n…[truncated, ${text.length} characters: ask for less]` : text;
        } catch (error) {
          if (signal.aborted) throw error;
          step.error = true;
          step.result = error instanceof z.ZodError ? `Invalid arguments: ${z.prettifyError(error)}` : error instanceof Error ? error.message : String(error);
        }
        step.ms = Math.round(performance.now() - started);
        results.push(tag("result", step.result || "(done)", `tool="${call.name}" status="${step.error ? "error" : "ok"}"`));
        yield { steps, llm };
      }
      const now = await agent.context().catch(() => context);
      prompt = [
        tag("tool_results", results.join("\n")),
        answer && tag("note", `"answer" was ignored: it must be alone in its reply.`),
        now !== context && tag("context_now", now),
        FORMAT,
      ]
        .filter(Boolean)
        .join("\n\n");
      context = now;
    }
  } finally {
    await session.close();
  }
}
