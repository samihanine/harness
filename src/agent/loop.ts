/**
 * The agent of every screen, for a small plain-text LLM: each reply is one JSON object listing tool
 * calls (`{"tools":[{"name","args"}]}`), parsed tolerantly; the application runs them and sends back
 * the results. A user message goes through two roles, each with its own prompt and tools:
 *
 * 1. The PLANNER (one LLM conversation per app conversation, read tools only) explores, then answers
 *    at once when nothing has to change, or asks the user to choose (`ask_user`: questions with options, shown as buttons in the chat) with
 *    the action plan it would follow. When the user has answered, it either asks again, answers
 *    (nothing to change) or hands the final plan to the implementer (`implement`).
 * 2. The IMPLEMENTER (a fresh LLM conversation per plan, dropped afterwards, every tool) does the plan
 *    and answers; the screen's `finish` then saves its changes (a refused save goes back to it once).
 *
 * Every message ends with the reply format and the tool names: weak models forget both.
 */
import { z } from "zod";
import { AiAuthError, deleteRemoteConversation, openSession, type Session } from "./llm";
import { ReplyError, parseReply } from "./reply";
import { LEARN_FROM_MISTAKES } from "@/config";
import { addMistakes, mistakesText, type Mistake } from "@/lib/mistakes";
import { settings } from "@/lib/store";
import { DIRECT, IMPLEMENTER, NOTES, PLANNER, PROTOCOL } from "@/prompts/agent";

export type Tool = {
  name: string;
  description: string;
  args: z.ZodObject;
  /** Reads only: given to the planner too (safe to run before the user validated). */
  readOnly?: boolean;
  run: (args: any) => Promise<string>;
};

export const tool = <S extends z.ZodObject>(t: { name: string; description: string; args: S; readOnly?: boolean; run: (args: z.infer<S>) => Promise<string> }): Tool => t;

export type Agent = {
  /** What this screen is, what the AI can do there and how (shared by the planner and the implementer). */
  instructions: string;
  tools: Tool[];
  /**
   * What the AI works on (model, file, notes…): sent at the start of a conversation, again only when it
   * changes. `focus`: the user's request (a big model then details only the tables it is about).
   */
  context: (focus?: string) => Promise<string>;
  /** What the user sees now (page, filters, rows…): re-read before each model call, sent when it changed. */
  state?: (progress?: Progress) => Promise<string>;
  /**
   * Its actions are harmless and undone in a click (filters, slicers, pages of a viewed report): no
   * planning, no question — one agent acts at once on the user's message.
   */
  direct?: boolean;
  /**
   * Its answers are changes the user expects without being asked (e.g. the Data Scientist's canvas): the
   * planner hands a clear request to the implementer at once, and asks only when it is unclear.
   */
  autoImplement?: boolean;
  /** Its answers are written on the screen (e.g. the Data Scientist's canvas): a chat answer only when the user asks for one. */
  answerOnScreen?: boolean;
  /** Run when the implementer has finished (e.g. saving its draft); its text is added to the answer. */
  finish?: () => Promise<string | undefined>;
};

export type Step = { id: string; name: string; args: Record<string, unknown>; result?: string; error?: boolean; ms?: number };
/** One exchange with the model, exactly as sent and received. */
export type LlmCall = { at: string; role: "planner" | "implementer"; prompt: string; reply?: string; error?: string; ms?: number };
/** Reports how far a long read is (e.g. visuals of the page read so far / all of them). */
export type Progress = (done: number, total: number) => void;
/** What the agent is doing between visible steps, with progress when known. */
export type Status = { label: string; done?: number; total?: number };
export type Update = { steps: Step[]; llm: LlmCall[]; answer?: string; remote?: string; status?: Status };

/** The planner's questions, as shown in the chat (an `ask_user` step). */
export type Choices = { message?: string; questions: { question: string; options: string[]; multiple?: boolean }[]; plan: string };

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

/** sessionChars: past this size the planner's conversation is restarted (with a summary of the app conversation). */
/** emptyRetries: clean sessions tried when the first reply is empty (the model returns nothing on ~1 call in 3, at random). */
const LIMITS = { retries: 2, refusals: 3, sendRetries: 2, emptyRetries: 4, resultChars: 6_000, historyChars: 12_000, sessionChars: 150_000 };

/** Tool turns of the implementer (Settings › AI › maximum steps); the planner gets half, at least 6. */
const stepsOf = async () => {
  const max = Math.max(4, Math.min(100, Number((await settings.get()).maxSteps) || 24));
  return { implementer: max, planner: Math.max(6, Math.round(max / 2)) };
};

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
  description: "Shows the user your findings and questions with options to click (they can also answer freely), with your action plan. Ends your turn. Alone in its reply.",
  args: z.object({
    message: z.string().optional().describe("Markdown for the user: what you found / understood, or the answer to their question"),
    questions: z.array(z.object({ question: z.string(), options: z.array(option).min(2).max(5).describe("Short answers, the most likely FIRST"), multiple: z.boolean().optional().describe("Several options may be chosen") })).min(1).max(4),
    plan: z.string().describe("The action plan for the implementer if the user agrees: numbered steps, exact tools, names and values (not shown)"),
  }),
  readOnly: true,
  run: async () => "",
});
const IMPLEMENT = tool({
  name: "implement",
  description: "Hands the final plan to the implementer, who makes the changes now. Only after the user answered your questions. Alone in its reply.",
  args: z.object({ plan: z.string().describe("Self-contained final plan: every step with exact tools, names and values, and the user's choices") }),
  readOnly: true,
  run: async () => "",
});

/** End of EVERY message: the format, the tools by kind, and what is expected now. */
const format = (tools: Tool[], now: string, example: string) =>
  tag(
    "reply_format",
    `Reply with ONE raw JSON object and NOTHING else: {"tools": [{"name": "<tool>", "args": {…}}]}
Read tools: ${tools.filter((t) => t.readOnly && !TERMINALS.has(t.name)).map((t) => t.name).join(", ") || "none"}.
${tools.some((t) => !t.readOnly) ? `Tools that change things: ${tools.filter((t) => !t.readOnly).map((t) => t.name).join(", ")}.\n` : ""}Final tools (alone in their reply): ${tools.filter((t) => TERMINALS.has(t.name)).map((t) => t.name).join(", ")}.
Now: ${now}
Example: ${example}`,
  );
const TERMINALS = new Set(["answer", "ask_user", "implement"]);
/** A request asking for an answer in the chat only (autoImplement screens otherwise answer on the screen). */
const CHAT_ONLY = /\b(in|dans) (the |le )?chat\b|\bjuste (me )?(dire|répondre)|\bjust (tell|answer)\b/i;

/** The planner's live conversation per app conversation, with what it was last told. */
type Live = { session: Session; model: string; context?: string; state?: string; note?: string };
const live = new Map<string, Live>();
/** Drops the planner conversation of an app conversation (deleted). */
export const forgetSession = (conversationId: string) => void live.delete(conversationId);

/** `invalid`: replies that could not be read (kept as mistakes). */
type Turn = { steps: Step[]; llm: LlmCall[]; remote?: string; invalid?: string[] };
/** How a role ended: its final tool, or a message for the user when it could not finish. */
type Ended = { name: string; args: Record<string, unknown> } | { giveUp: string };
type Seen = { context?: string; state?: string };

/**
 * Runs one role until it calls one of its final tools (returned) or gives up (a message for the user).
 * Tools run in order; after them the page state is read again and sent when it changed.
 */
async function* drive(opts: {
  role: LlmCall["role"];
  session: Session;
  first: string;
  tools: Tool[];
  max: number;
  turn: Turn;
  signal: AbortSignal;
  /** Re-reads context and page state; returns the tags of what changed. */
  reread: () => AsyncGenerator<Update, string[]>;
  /** End of every message (format and what is expected). */
  tail: () => string;
  /** A final tool refused now (e.g. implement before the user answered): the reason sent back. */
  refuse?: (name: string, args: Record<string, unknown>) => string | undefined;
  /**
   * A new session, when this one has nothing yet but `first`: an empty first reply (the model sometimes
   * returns nothing on a long prompt) is asked again in a clean session, not after the empty answer.
   */
  fresh?: () => Session;
}): AsyncGenerator<Update, Ended> {
  const { turn, signal } = opts;
  let session = opts.session;
  let exchanges = 0;
  let empties = 0;
  const byName = new Map(opts.tools.map((t) => [t.name, t]));
  let prompt = opts.first;
  let retries = 0;
  let refusals = 0;
  let reminded = false;
  for (let n = 0; n <= opts.max; n++) {
    signal.throwIfAborted();
    const call: LlmCall = { at: new Date().toISOString(), role: opts.role, prompt: n === opts.max ? `${tag("note", NOTES.lastTurn)}\n\n${opts.tail()}` : prompt };
    turn.llm.push(call);
    yield { ...turn, status: { label: opts.role === "planner" ? (turn.steps.length ? "Exploring" : "Understanding your request") : "Working on the plan" } };
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
      session = opts.fresh();
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
        // Kept in the trace (it used to be invisible), and not counted as a step a few times.
        (turn.invalid ??= []).push(`refused ${final.name}: ${problem.slice(0, 300)}`);
        if (++refusals <= LIMITS.refusals) n--;
        prompt = `${tag("refused", problem)}\n\n${opts.tail()}`;
        continue;
      }
      // Failed steps never redone: one reminder, so the model retries them or says what failed.
      const failed = turn.steps.filter((s, i) => s.error && !turn.steps.slice(i + 1).some((t) => t.name === s.name && !t.error));
      if (opts.role === "implementer" && failed.length && !reminded) {
        reminded = true;
        prompt = `${tag("note", `Before answering: these steps FAILED and were not redone successfully:\n${failed.map((s) => `- ${s.name} ${JSON.stringify(s.args).slice(0, 200)} → ${String(s.result).slice(0, 300)}`).join("\n")}\nFix and retry them now, or say clearly in your answer what is not done.`)}\n\n${opts.tail()}`;
        continue;
      }
      return { name: final.name, args: parsed.data as Record<string, unknown> };
    }
    const results: string[] = [];
    for (const c of calls) {
      if (TERMINALS.has(c.name)) {
        results.push(tag("result", `"${c.name}" was ignored: ${NOTES.aloneFinal}`, `tool="${c.name}" status="error"`));
        continue;
      }
      const step: Step = { id: crypto.randomUUID(), name: c.name, args: c.args };
      turn.steps.push(step);
      yield { ...turn };
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
      results.push(tag("result", step.result || "(done)", `tool="${c.name}" status="${step.error ? "error" : "ok"}"`));
      yield { ...turn };
    }
    prompt = [tag("tool_results", results.join("\n")), ...(yield* opts.reread()), opts.tail()].filter(Boolean).join("\n\n");
  }
  return { giveUp: `⚠️ The assistant used all its ${opts.max} steps without finishing (Settings › AI › maximum steps). Ask again, more precisely.` };
}

/** The user's answers to the questions, as sent back to the planner. */
export const answersText = (choices: Choices, answers: string[][]) => choices.questions.map((q, i) => `- ${q.question} → ${answers[i]?.join(", ") || "(no answer)"}`).join("\n");

/** The turn's failed calls, each with the call of the same tool that worked after it (if any). */
function mistakesOf(turn: Turn, scope: string): Mistake[] {
  const at = new Date().toISOString();
  const args = (s: Step) => JSON.stringify(s.args).slice(0, 400);
  return [
    ...turn.steps.flatMap((s, i) => {
      if (!s.error) return [];
      const fix = turn.steps.slice(i + 1).find((t) => t.name === s.name && !t.error);
      return [{ at, scope, tool: s.name, args: args(s), error: String(s.result ?? "").slice(0, 500), fix: fix && args(fix) }];
    }),
    ...(turn.invalid ?? []).map((error) => ({ at, scope, tool: "(reply)", args: "", error })),
  ];
}

type TurnArgs = Parameters<typeof runTurn>[0];
/** The request each app conversation is about (kept while the user answers the planner's questions). */
const focusOf = new Map<string, string>();

/** One user message: planner then implementer (see the top of this file); the mistakes made are kept. */
export async function* runTurn(args: {
  agent: Agent;
  /** The planner's model. */
  model: string;
  /** The implementer's model (Settings › AI: often a different one than the planner's). */
  implementerModel?: string;
  /** App conversation: its planner conversation is reused from one message to the next. */
  conversationId: string;
  /** The conversation so far, as text: only used to start a new planner conversation (e.g. after a reload). */
  history: string;
  request: string;
  /** The message answers the planner's last questions (the planner may then implement or answer). */
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

async function* turnOf({ agent, model, implementerModel = model, conversationId, history, request, answering, signal }: TurnArgs, turn: Turn): AsyncGenerator<Update> {
  // Past mistakes of this screen's tools: shown once to each new conversation, as examples not to follow.
  const mistakes = LEARN_FROM_MISTAKES ? await mistakesText(agent.tools.map((t) => t.name)) : "";
  const mistakesTag = mistakes && tag("past_mistakes", `Calls that failed before on this screen (do not repeat them; "worked instead" is the call that fixed it):
${mistakes}`);
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

  /* --------------------------------- Planner -------------------------------- */
  // Direct agents skip it: the implementer acts on the message itself.
  let ended: Ended = { name: "implement", args: { plan: "Do the user's last message (in <conversation>), directly." } };
  if (!agent.direct) {
    const readTools = [...agent.tools.filter((t) => t.readOnly), ASK_USER, IMPLEMENT, ANSWER];
    let conv = live.get(conversationId);
    if (conv && (conv.model !== model || conv.session.chars > LIMITS.sessionChars)) conv = undefined;
    const fresh = !conv;
    if (!conv) live.set(conversationId, (conv = { session: openSession(model, "planner", conversationId), model }));
    const planner = conv;
    const now = yield* readShown("Reading what you see");
    const plannerTail = (stage: "new" | "answers" | "exploring") =>
      format(
        readTools,
        stage === "answers"
          ? NOTES.plannerAnswers
          : stage === "new" && agent.autoImplement
            ? agent.answerOnScreen
              ? NOTES.plannerNewOnScreen
              : NOTES.plannerNewAuto
            : stage === "new"
            ? NOTES.plannerNew
            : NOTES.plannerExploring,
        stage === "answers" ? `{"tools": [{"name": "implement", "args": {"plan": "1. … 2. …"}}]}` : `{"tools": [{"name": "ask_user", "args": {"message": "…", "questions": [{"question": "…?", "options": ["Yes, do it", "No"]}], "plan": "1. …"}}]}`,
      );
    const first = [
      ...(fresh ? [tag("protocol", PROTOCOL), tag("role", PLANNER), tag("app_instructions", agent.instructions), tag("tools", readTools.map(signature).join("\n")), mistakesTag] : []),
      ...changes(planner, now, fresh),
      fresh && history && tag("conversation_so_far", history.slice(-LIMITS.historyChars)),
      planner.note && tag("since_your_last_message", planner.note),
      answering ? tag("user_answers", request) : tag("user_message", request),
      plannerTail(answering ? "answers" : "new"),
    ]
      .filter(Boolean)
      .join("\n\n");
    planner.note = undefined;
    let plannerStage: "answers" | "exploring" = answering ? "answers" : "exploring";
    try {
      ended = yield* drive({
        role: "planner",
        session: planner.session,
        // A new conversation has only this message: an empty reply is asked again in a clean one.
        fresh: fresh ? () => (planner.session = openSession(model, "planner", conversationId)) : undefined,
        first,
        tools: readTools,
        max: (await stepsOf()).planner,
        turn,
        signal,
        reread: async function* () {
          return changes(planner, yield* readShown("Reading the page again"));
        },
        tail: () => plannerTail(plannerStage),
        // Changes need the user's agreement; a plain answer does not.
        refuse: (name) =>
          name === "implement" && !answering && !agent.autoImplement
            ? NOTES.notAgreed
            : // Its answers are written on the screen (e.g. the canvas): the chat only when the user asks for it.
              name === "answer" && agent.answerOnScreen && !CHAT_ONLY.test(request)
              ? NOTES.answerOnScreen
              : undefined,
      });
    } catch (error) {
      // A failed exchange may have left the conversation incomplete: the next message starts a new one.
      if (live.get(conversationId) === planner) live.delete(conversationId);
      throw error;
    }
    plannerStage = "exploring";
    if ("giveUp" in ended) return yield { ...turn, answer: ended.giveUp };
    if (ended.name === "ask_user") {
      turn.steps.push({ id: crypto.randomUUID(), name: "ask_user", args: ended.args, result: "" });
      return yield { ...turn, answer: "" };
    }
    if (ended.name === "answer") return yield { ...turn, answer: String(ended.args.text ?? "") };
  }

  /* ------------------------------- Implementer ------------------------------ */
  const plan = String(ended.args.plan ?? "");
  if (!agent.direct) turn.steps.push({ id: crypto.randomUUID(), name: "implement", args: { plan }, result: "" });
  yield { ...turn };
  let session = openSession(model, "implementer", conversationId);
  const seen: Seen = {};
  const tools = [...agent.tools, ANSWER];
  const implementerTail = () => format(tools, NOTES.implementerNext, `{"tools": [{"name": "${agent.tools.find((t) => !t.readOnly)?.name ?? "answer"}", "args": {…}}]}`);
  const start = [
    tag("protocol", PROTOCOL),
    tag("role", agent.direct ? DIRECT : IMPLEMENTER),
    tag("app_instructions", agent.instructions),
    tag("tools", tools.map(signature).join("\n")),
    mistakesTag,
    ...changes(seen, yield* readShown("Reading what you see")),
    tag("conversation", `${history.slice(-3000)}\n\nuser: ${request}`.trim()),
    tag("plan", plan),
    implementerTail(),
  ]
    .filter(Boolean)
    .join("\n\n");
  const implemented = turn.steps.length;
  let result: Ended = { giveUp: "" };
  let saveRetried = false;
  let note: string | undefined;
  try {
    for (;;) {
      result = yield* drive({
        role: "implementer",
        session,
        first: saveRetried ? `${tag("save_failed", note ?? "")}\n\n${implementerTail()}` : start,
        fresh: saveRetried ? undefined : () => (session = openSession(model, "implementer", conversationId)),
        tools,
        max: (await stepsOf()).implementer,
        turn,
        signal,
        reread: async function* () {
          return changes(seen, yield* readShown("Reading the page again"));
        },
        tail: implementerTail,
      });
      if ("giveUp" in result || !agent.finish) break;
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
  } finally {
    void session.remote().then((id) => {
      if (id) void deleteRemoteConversation(id).catch(() => undefined);
    });
  }
  const text = "giveUp" in result ? result.giveUp : String(result.args.text ?? "");
  // The planner learns what was done before the user's next message.
  const done = turn.steps.slice(implemented).map((s) => `- ${s.name}${s.error ? " (failed)" : ""}: ${String(s.result ?? "").split("\n")[0].slice(0, 160)}`);
  const conv = live.get(conversationId);
  if (conv) conv.note = `The implementer carried out your plan.\nActions:\n${done.join("\n") || "none"}\nIts answer to the user: ${text}${note ? `\n${note}` : ""}`;
  yield { ...turn, answer: note ? `${text}\n\n${note}` : text };
}
