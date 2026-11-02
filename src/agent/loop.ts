/**
 * The agent of every screen, for a small plain-text LLM: each reply is one JSON object listing tool
 * calls (`{"tools":[{"name","args"}]}`), parsed tolerantly; the application runs them and sends back
 * the results. One LLM conversation per app conversation, kept from one message to the next (the
 * context is sent once, then only what changed), with every tool of the screen:
 *
 * - nothing to change (a question, figures): it reads what it needs and answers;
 * - changes: it first shows what it will do and asks the user to confirm (`ask_user`: questions with
 *   options, shown as buttons in the chat), then acts on the answer. Tools that change things are
 *   refused before the user confirmed, unless the request is obvious: the model then says so with
 *   `proceed` in the same reply. Screens whose actions are undone in a click (`direct`) or expected
 *   without asking (`autoImplement`) act at once.
 *
 * The screen's `finish` saves the changes when the answer is done (a refused save goes back once).
 * Every message ends with the reply format: weak models forget it.
 */
import { z } from "zod";
import { AiAuthError, openSession, type Session } from "./llm";
import { ReplyError, parseReply } from "./reply";
import { LEARN_FROM_MISTAKES } from "@/config";
import { contextRound } from "@/features/confluence/context-pages";
import { addMistakes, alike, mistakesText, type Mistake } from "@/lib/mistakes";
import { settings } from "@/lib/store";
import { AGENT, DIRECT, NOTES, PROTOCOL } from "@/prompts/agent";

export type Tool = {
  name: string;
  description: string;
  args: z.ZodObject;
  /** Reads only: runs before the user confirmed anything. */
  readOnly?: boolean;
  run: (args: any) => Promise<string>;
};

export const tool = <S extends z.ZodObject>(t: { name: string; description: string; args: S; readOnly?: boolean; run: (args: z.infer<S>) => Promise<string> }): Tool => t;

export type Agent = {
  /** What this screen is, what the AI can do there and how. */
  instructions: string;
  tools: Tool[];
  /**
   * What the AI works on (model, file, notes…): sent at the start of a conversation, again only when it
   * changes. `focus`: the user's request (a big model then details only the tables it is about).
   */
  context: (focus?: string) => Promise<string>;
  /** What the user sees now (page, filters, rows…): re-read before each model call, sent when it changed. */
  state?: (progress?: Progress) => Promise<string>;
  /** Its actions are harmless and undone in a click (filters, slicers, pages of a viewed report): no question, it acts at once. */
  direct?: boolean;
  /** Its answers are changes the user expects without being asked (e.g. the Data Scientist's canvas): no confirmation, a question only when the request is unclear. */
  autoImplement?: boolean;
  /** Its answers are written on the screen (e.g. the Data Scientist's canvas): a chat answer only when the user asks for one. */
  answerOnScreen?: boolean;
  /** What is still missing before the final answer (e.g. the Data Scientist's pivot), sent back to the model; nothing: it may answer. */
  beforeAnswer?: () => string | undefined;
  /** Run when the answer is done and something changed (e.g. saving the draft); its text is added to the answer. */
  finish?: () => Promise<string | undefined>;
};

export type Step = { id: string; name: string; args: Record<string, unknown>; result?: string; error?: boolean; ms?: number };
/** One exchange with the model, exactly as sent and received. */
export type LlmCall = { at: string; role: "agent"; prompt: string; reply?: string; error?: string; ms?: number };
/** Reports how far a long read is (e.g. visuals of the page read so far / all of them). */
export type Progress = (done: number, total: number) => void;
/** What the agent is doing between visible steps, with progress when known. */
export type Status = { label: string; done?: number; total?: number };
export type Update = { steps: Step[]; llm: LlmCall[]; answer?: string; remote?: string; status?: Status };

/** The questions asked before acting, as shown in the chat (an `ask_user` step); `plan`: older messages. */
export type Choices = { message?: string; questions: { question: string; options: string[]; multiple?: boolean }[]; plan?: string };

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

/**
 * sessionChars: past this size the conversation is restarted (with the app conversation as text): a long
 * one makes every call slow. emptyRetries: clean sessions tried when the first reply is empty (the model
 * returns nothing on some calls, at random).
 */
const LIMITS = { retries: 2, refusals: 3, sendRetries: 2, emptyRetries: 4, resultChars: 4_000, historyChars: 8_000, sessionChars: 60_000 };

/** Tool turns per message (Settings › AI › maximum steps). */
const stepsOf = async () => Math.max(4, Math.min(100, Number((await settings.get()).maxSteps) || 24));

/** Failures worth sending the message again: the network, the API's passing errors (not a refused key). */
const TRANSIENT = /failed to fetch|network|load failed|invalid completion|timed? ?out|\b(429|500|502|503|504)\b|ECONN|socket/i;

/** Sends a message, again after a short wait when the network or the API failed for a moment. */
async function sendRetrying(session: Session, text: string, signal: AbortSignal) {
  for (let attempt = 0; ; attempt++) {
    try {
      return await session.send(text, signal);
    } catch (error) {
      if (signal.aborted || error instanceof AiAuthError || attempt >= LIMITS.sendRetries || !TRANSIENT.test(describeError(error))) throw error;
      await new Promise((r) => setTimeout(r, 1500 * (attempt + 1) ** 2));
    }
  }
}

const tag = (name: string, body: string, attrs = "") => `<${name}${attrs ? ` ${attrs}` : ""}>\n${body.trim()}\n</${name}>`;

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

const ANSWER = tool({ name: "answer", description: "Final reply to the user (markdown). Ends the turn. Alone in its reply.", args: z.object({ text: z.string() }), readOnly: true, run: async () => "" });
const option = z.preprocess((o) => (o && typeof o === "object" && "label" in o ? (o as { label: unknown }).label : o), z.string());
const ASK_USER = tool({
  name: "ask_user",
  description: "Before changing anything: shows the user what you found and what you will do, with questions to click (they can also answer freely). Ends your turn. Alone in its reply.",
  args: z.object({
    message: z.string().optional().describe("Markdown for the user: what you understood and exactly what you will change"),
    questions: z.array(z.object({ question: z.string(), options: z.array(option).min(2).max(5).describe("Short answers, the most likely FIRST"), multiple: z.boolean().optional().describe("Several options may be chosen") })).min(1).max(4),
  }),
  readOnly: true,
  run: async () => "",
});
const PROCEED = tool({
  name: "proceed",
  description: "Only when the request is obvious (every name and value given, one way to do it): put it in the same reply as the tools that change things, which then run without asking the user.",
  args: z.object({ reason: z.string().describe("Why nothing needs confirming, in a few words") }),
  readOnly: true,
  run: async () => "Acting without asking.",
});

/** End of EVERY message: the format and what is expected now (short: it is sent each time). */
const format = (tools: Tool[], now: string) =>
  tag(
    "reply_format",
    `ONE raw JSON object, nothing else: {"tools": [{"name": "<tool>", "args": {…}}]}. Tools: ${tools.map((t) => t.name).join(", ")}. Final tools (alone in their reply): ${tools.filter((t) => TERMINALS.has(t.name)).map((t) => t.name).join(", ")}.
Now: ${now}`,
  );
const TERMINALS = new Set(["answer", "ask_user"]);
/** The tools that read a model's data: a request is explored with them before the user is asked. */
const DATA_READS = new Set(["run_dax", "query"]);
/** A request asking for an answer in the chat only (autoImplement screens otherwise answer on the screen). */
const CHAT_ONLY = /\b(in|dans) (the |le )?chat\b|\bjuste (me )?(dire|répondre)|\bjust (tell|answer)\b/i;

/** The live conversation of an app conversation, with what it was last told. */
type Live = { session: Session; model: string; context?: string; state?: string };
const live = new Map<string, Live>();
/** Drops the model conversation of an app conversation (deleted). */
export const forgetSession = (conversationId: string) => void live.delete(conversationId);

/** `invalid`: replies that could not be read (kept as mistakes). */
type Turn = { steps: Step[]; llm: LlmCall[]; remote?: string; invalid?: string[] };
/** How a turn ended: its final tool, or a message for the user when it could not finish. */
type Ended = { name: string; args: Record<string, unknown> } | { giveUp: string };
type Seen = { context?: string; state?: string };

/**
 * Runs the conversation until the model calls a final tool (returned) or gives up (a message for the
 * user). Tools run in order; after them the page state is read again and sent when it changed.
 */
async function* drive(opts: {
  session: () => Session;
  first: string;
  tools: Tool[];
  max: number;
  turn: Turn;
  signal: AbortSignal;
  /** Re-reads context and page state; returns the tags of what changed. */
  reread: () => AsyncGenerator<Update, string[]>;
  /** End of every message (format and what is expected). */
  tail: () => string;
  /** Tools that change things refused now (the user has not confirmed): the reason sent back. */
  gate: (calls: { name: string }[]) => string | undefined;
  /** A final tool refused now: the reason sent back. */
  refuse?: (name: string, args: Record<string, unknown>) => string | undefined;
  /**
   * A new session, when this one has nothing yet but `first`: an empty first reply (the model sometimes
   * returns nothing on a long prompt) is asked again in a clean session, not after the empty answer.
   */
  fresh?: () => void;
}): AsyncGenerator<Update, Ended> {
  const { turn, signal } = opts;
  let exchanges = 0;
  let empties = 0;
  const byName = new Map(opts.tools.map((t) => [t.name, t]));
  let prompt = opts.first;
  let retries = 0;
  let refusals = 0;
  let reminded = false;
  for (let n = 0; n <= opts.max; n++) {
    signal.throwIfAborted();
    const session = opts.session();
    const call: LlmCall = { at: new Date().toISOString(), role: "agent", prompt: n === opts.max ? `${tag("note", NOTES.lastTurn)}\n\n${opts.tail()}` : prompt };
    turn.llm.push(call);
    const changed = turn.steps.some((s) => !s.error && !byName.get(s.name)?.readOnly);
    yield { ...turn, status: { label: changed ? "Working" : turn.steps.length ? "Exploring" : "Understanding your request" } };
    const started = performance.now();
    exchanges++;
    const reply = await sendRetrying(session, call.prompt, signal).catch((error: unknown) => {
      call.error = describeError(error);
      throw error;
    });
    Object.assign(call, { reply, ms: Math.round(performance.now() - started) });
    turn.remote ??= await session.remote();
    if (!reply.trim() && exchanges === 1 && opts.fresh && empties < LIMITS.emptyRetries) {
      empties++;
      opts.fresh();
      exchanges = 0;
      n--;
      continue;
    }
    let calls;
    try {
      calls = parseReply(reply);
      const unknown = calls.filter((c) => !byName.has(c.name)).map((c) => c.name);
      if (unknown.length) throw new ReplyError(`Unknown tool(s): ${unknown.join(", ")}. Your tools: ${[...byName.keys()].join(", ")}.`);
      retries = 0;
    } catch (error) {
      if (!(error instanceof ReplyError)) throw error;
      // Plain prose: the model talking to the user. Sent back once to get the proper final tool.
      (turn.invalid ??= []).push(`${error.message} — reply: ${reply.slice(0, 300)}`);
      if (++retries > LIMITS.retries) return { giveUp: !reply.includes("{") && reply.trim() ? reply.trim() : `⚠️ The assistant's reply could not be read (${error.message}). Ask again.` };
      prompt = `${tag("invalid_reply", `Your reply could not be used: ${error.message}\n${NOTES.invalidReply}`)}\n\n${opts.tail()}`;
      n--;
      continue;
    }
    const final = calls.find((c) => TERMINALS.has(c.name));
    if (final && calls.length === 1) {
      const refused = opts.refuse?.(final.name, final.args);
      // A final tool with wrong arguments: refused with the reason, like any tool.
      const parsed = byName.get(final.name)!.args.safeParse(final.args);
      const problem = refused ?? (parsed.success ? undefined : `Invalid arguments of ${final.name}: ${z.prettifyError(parsed.error)}`);
      if (problem) {
        (turn.invalid ??= []).push(`refused ${final.name}: ${problem.slice(0, 300)}`);
        if (++refusals <= LIMITS.refusals) n--;
        prompt = `${tag("refused", problem)}\n\n${opts.tail()}`;
        continue;
      }
      // Failed steps never redone: one reminder, so the model retries them or says what failed.
      const failed = turn.steps.filter((s, i) => s.error && !turn.steps.slice(i + 1).some((t) => t.name === s.name && !t.error));
      if (final.name === "answer" && failed.some((s) => !byName.get(s.name)?.readOnly) && !reminded) {
        reminded = true;
        prompt = `${tag("note", `Before answering: these steps FAILED and were not redone successfully:\n${failed.map((s) => `- ${s.name} ${JSON.stringify(s.args).slice(0, 200)} → ${String(s.result).slice(0, 300)}`).join("\n")}\nFix and retry them now, or say clearly in your answer what is not done.`)}\n\n${opts.tail()}`;
        continue;
      }
      return { name: final.name, args: parsed.data as Record<string, unknown> };
    }
    const gated = opts.gate(calls);
    const results: string[] = [];
    const run = async (c: { name: string; args: Record<string, unknown> }, step: Step) => {
      const t = byName.get(c.name)!;
      const t0 = performance.now();
      try {
        const text = await t.run(t.args.parse(c.args));
        step.result = text.length > LIMITS.resultChars ? `${text.slice(0, LIMITS.resultChars)}\n…[truncated, ${text.length} characters: ask for less]` : text;
      } catch (error) {
        if (signal.aborted) throw error;
        step.error = true;
        step.result = error instanceof z.ZodError ? `Invalid arguments: ${z.prettifyError(error)}` : describeError(error);
      }
      step.ms = Math.round(performance.now() - t0);
      return tag("result", step.result || "(done)", `tool="${c.name}" status="${step.error ? "error" : "ok"}"`);
    };
    for (let i = 0; i < calls.length; i++) {
      const c = calls[i];
      if (TERMINALS.has(c.name)) {
        results.push(tag("result", `"${c.name}" was ignored: ${NOTES.aloneFinal}`, `tool="${c.name}" status="error"`));
        continue;
      }
      const t = byName.get(c.name)!;
      if (gated && !t.readOnly) {
        results.push(tag("result", gated, `tool="${c.name}" status="refused"`));
        continue;
      }
      // Reads next to each other (several DAX queries, say) run at once; changes keep their order.
      let end = i + 1;
      if (t.readOnly) while (end < calls.length && !TERMINALS.has(calls[end].name) && byName.get(calls[end].name)?.readOnly) end++;
      const batch = calls.slice(i, end).map((call) => ({ call, step: { id: crypto.randomUUID(), name: call.name, args: call.args } as Step }));
      turn.steps.push(...batch.map((b) => b.step));
      yield { ...turn };
      results.push(...(await Promise.all(batch.map((b) => run(b.call, b.step)))));
      yield { ...turn };
      i = end - 1;
    }
    prompt = [tag("tool_results", results.join("\n")), ...(yield* opts.reread()), opts.tail()].filter(Boolean).join("\n\n");
  }
  return { giveUp: `⚠️ The assistant used all its ${opts.max} steps without finishing (Settings › AI › maximum steps). Ask again, more precisely.` };
}

/** The user's answers to the questions, as sent back to the model. */
export const answersText = (choices: Choices, answers: string[][]) => choices.questions.map((q, i) => `- ${q.question} → ${answers[i]?.join(", ") || "(no answer)"}`).join("\n");

/** The turn's failed calls, each with the call of the same tool that worked after it (if any). */
function mistakesOf(turn: Turn, scope: string): Mistake[] {
  const at = new Date().toISOString();
  const args = (s: Step) => JSON.stringify(s.args).slice(0, 400);
  return [
    ...turn.steps.flatMap((s, i) => {
      if (!s.error) return [];
      const fix = turn.steps.slice(i + 1).find((t) => t.name === s.name && !t.error && alike(args(s), args(t)));
      return [{ at, scope, tool: s.name, args: args(s), error: String(s.result ?? "").slice(0, 500), fix: fix && args(fix) }];
    }),
    ...(turn.invalid ?? []).map((error) => ({ at, scope, tool: "(reply)", args: "", error })),
  ];
}

type TurnArgs = Parameters<typeof runTurn>[0];
/** The request each app conversation is about (kept while the user answers the questions). */
const focusOf = new Map<string, string>();

/** One user message (see the top of this file); the mistakes made are kept. */
export async function* runTurn(args: {
  agent: Agent;
  model: string;
  /** App conversation: its model conversation is reused from one message to the next. */
  conversationId: string;
  /** The conversation so far, as text: only used to start a new model conversation (e.g. after a reload). */
  history: string;
  request: string;
  /** The message answers the last questions: the user confirmed (or redirected) the change. */
  answering: boolean;
  signal: AbortSignal;
  /** The screen (kept with the mistakes). */
  scope?: string;
}): AsyncGenerator<Update> {
  const turn: Turn = { steps: [], llm: [] };
  try {
    yield* turnOf(args, turn);
  } finally {
    void addMistakes(mistakesOf(turn, args.scope ?? ""));
  }
}

async function* turnOf({ agent, model, conversationId, history, request, answering, signal, scope }: TurnArgs, turn: Turn): AsyncGenerator<Update> {
  contextRound(conversationId);
  if (!answering || !focusOf.has(conversationId)) focusOf.set(conversationId, request);
  const focus = focusOf.get(conversationId);
  const read = async (progress?: Progress) => ({
    context: await agent.context(focus).catch((e) => `Context unavailable: ${describeError(e)}`),
    state: agent.state ? await agent.state(progress).catch((e) => `Page state unavailable: ${describeError(e)}`) : undefined,
  });
  /** `read`, telling the user what is happening (with progress) when it takes more than a moment. */
  const readShown = async function* (label: string): AsyncGenerator<Update, Awaited<ReturnType<typeof read>>> {
    let progress: { done: number; total: number } | undefined;
    let result: Awaited<ReturnType<typeof read>> | undefined;
    const task = read((done, total) => (progress = { done, total })).then((r) => (result = r));
    for (;;) {
      await Promise.race([task, new Promise((r) => setTimeout(r, 200))]);
      if (result) return result;
      yield { ...turn, status: { label, ...progress } };
    }
  };
  /** What changed since the model was last told (`seen`), as tags (everything on a new conversation). */
  const changes = (seen: Seen, next: Awaited<ReturnType<typeof read>>, first = false) => {
    const parts = [
      (first || next.context !== seen.context) && tag(first ? "context" : "context_now", next.context),
      next.state !== undefined && (first || next.state !== seen.state) && tag(first ? "page_state" : "page_state_now", next.state),
    ].filter((p): p is string => !!p);
    seen.context = next.context;
    seen.state = next.state;
    return parts;
  };

  const asking = !agent.direct;
  // proceed only where changes wait for a confirmation (an agent acting at once has nothing to skip).
  const tools = [...agent.tools, ...(asking ? [...(agent.autoImplement ? [] : [PROCEED]), ASK_USER] : []), ANSWER];
  let conv = live.get(conversationId);
  if (conv && (conv.model !== model || conv.session.chars > LIMITS.sessionChars)) conv = undefined;
  const fresh = !conv;
  if (!conv) live.set(conversationId, (conv = { session: openSession(model, "agent", conversationId), model }));
  const current = conv;
  // Past mistakes of this screen's tools: shown once to each new conversation, as examples not to follow.
  const mistakes = fresh && LEARN_FROM_MISTAKES ? await mistakesText(agent.tools.map((t) => t.name), scope) : "";
  const now = yield* readShown("Reading what you see");
  // Changes need the user's agreement, unless the screen acts without asking or the user just answered.
  let confirmed = !asking || !!agent.autoImplement || answering;
  const stage = answering ? NOTES.answers : !asking ? NOTES.direct : agent.autoImplement ? (agent.answerOnScreen ? NOTES.newOnScreen : NOTES.newAuto) : NOTES.newRequest;
  let next = stage;
  const tail = () => format(tools, next);
  const first = [
    ...(fresh ? [tag("protocol", PROTOCOL), tag("role", asking ? AGENT : DIRECT), tag("app_instructions", agent.instructions), tag("tools", tools.map(signature).join("\n"))] : []),
    mistakes && tag("past_mistakes", `Calls that failed before on this screen (do not repeat them; "worked instead" is the call that fixed it):\n${mistakes}`),
    ...changes(current, now, fresh),
    fresh && history && tag("conversation_so_far", history.slice(-LIMITS.historyChars)),
    answering ? tag("user_answers", request) : tag("user_message", request),
    tail(),
  ]
    .filter(Boolean)
    .join("\n\n");
  const changedSteps = () => turn.steps.filter((s) => !s.error && !tools.find((t) => t.name === s.name)?.readOnly);
  let ended: Ended = { giveUp: "" };
  let note: string | undefined;
  let saveRetried = false;
  let explorePushed = false;
  try {
    for (;;) {
      ended = yield* drive({
        session: () => current.session,
        fresh: fresh && !saveRetried ? () => void (current.session = openSession(model, "agent", conversationId)) : undefined,
        first: saveRetried ? `${tag("save_failed", note ?? "")}\n\n${tail()}` : first,
        tools,
        max: await stepsOf(),
        turn,
        signal,
        reread: async function* () {
          next = changedSteps().length ? NOTES.working : stage;
          return changes(current, yield* readShown("Reading the page again"));
        },
        tail,
        gate: (calls) => {
          if (calls.some((c) => c.name === "proceed")) confirmed = true;
          return confirmed ? undefined : NOTES.notConfirmed;
        },
        refuse: (name) => {
          // On a screen with a model, a question to the user comes after exploring the data (sent back once).
          if (name === "ask_user" && !answering && !explorePushed && tools.some((t) => DATA_READS.has(t.name)) && !turn.steps.some((st) => DATA_READS.has(st.name))) {
            explorePushed = true;
            return NOTES.exploreFirst;
          }
          if (name !== "answer" || CHAT_ONLY.test(request)) return undefined;
          // Its answers are written on the screen (e.g. the canvas): the chat only when the user asks for it.
          if (agent.answerOnScreen && !changedSteps().length) return NOTES.answerOnScreen;
          return changedSteps().length ? agent.beforeAnswer?.() : undefined;
        },
      });
      if ("giveUp" in ended || ended.name !== "answer" || !agent.finish || !changedSteps().length) break;
      // The changes are saved before the answer is shown: a refused save goes back once (the draft is kept).
      yield { ...turn, status: { label: "Saving the changes" } };
      try {
        note = await agent.finish();
        break;
      } catch (error) {
        if (saveRetried) {
          note = `⚠️ Not saved: ${describeError(error)}`;
          break;
        }
        saveRetried = true;
        note = `Your changes could not be saved: ${describeError(error)}\nNothing was saved; your draft is kept. Fix the file named in the error (or discard_changes and redo the step), then answer.`;
      }
    }
  } catch (error) {
    // A failed exchange may have left the conversation incomplete: the next message starts a new one.
    if (live.get(conversationId) === current) live.delete(conversationId);
    throw error;
  }
  if ("giveUp" in ended) return yield { ...turn, answer: ended.giveUp };
  if (ended.name === "ask_user") {
    turn.steps.push({ id: crypto.randomUUID(), name: "ask_user", args: ended.args, result: "" });
    return yield { ...turn, answer: "" };
  }
  const text = String(ended.args.text ?? "");
  yield { ...turn, answer: note ? `${text}\n\n${note}` : text };
}
