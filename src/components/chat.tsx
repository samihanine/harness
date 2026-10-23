/**
 * AI chat (assistant-ui) on top of the planner / implementer agent (agent/loop.ts), with conversations
 * kept in IndexedDB. The planner's questions show as cards with options to click (or a free answer);
 * the answer goes back as the next user message.
 */
import { createContext, useContext, useEffect, useMemo, useRef, useState } from "react";
import {
  AssistantRuntimeProvider,
  ComposerPrimitive,
  MessagePrimitive,
  ThreadPrimitive,
  useAuiState,
  useLocalRuntime,
  type ChatModelAdapter,
  type ThreadMessage,
  type ToolCallMessagePartProps,
} from "@assistant-ui/react";
import { MarkdownTextPrimitive } from "@assistant-ui/react-markdown";
import remarkGfm from "remark-gfm";
import { ArrowUpIcon, CheckIcon, ChevronRightIcon, ChevronsUpDownIcon, DownloadIcon, FileTextIcon, ListChecksIcon, MessageCircleQuestionIcon, PaperclipIcon, PlusIcon, SendIcon, SquareIcon, Trash2Icon, WrenchIcon, XIcon } from "lucide-react";
import { Markdown as MarkdownView } from "./markdown";
import { Button } from "./ui/button";
import { Command, CommandEmpty, CommandGroup, CommandInput, CommandItem, CommandList } from "./ui/command";
import { Input } from "./ui/input";
import { Popover, PopoverContent, PopoverTrigger } from "./ui/popover";
import { forgetSession, runTurn, type Agent, type Choices, type Status, type Update } from "@/agent/loop";
import { AI_MODELS, addFileToConversation, deleteRemoteConversation, uploadAiFile } from "@/agent/llm";
import { PROMPTS, withPrompts, type PromptScreen } from "@/config";
import { find, put, remove, settings, useCollection, type Conversation } from "@/lib/store";
import { last } from "@/lib/last";
import { cleanMarkdown } from "./markdown";
import { cn } from "@/lib/utils";

/** Screen of a chat scope ("viewer:<id>", "excel:<id>"…), for the admin's additional instructions. */
const SCREEN_OF: Record<string, PromptScreen> = { viewer: "viewer", editor: "editor", research: "research", reports: "reports", datasets: "datasets", excel: "excel", contexts: "contexts", slides: "slides" };

/** The screen's agent with the instructions the admin added in Settings (for all screens, and this one). */
const withAdmin = (agent: Agent, scope: string): Agent => {
  const screen = SCREEN_OF[scope.split(":")[0]];
  return screen ? { ...agent, instructions: withPrompts(screen, agent.instructions) } : PROMPTS.all.trim() ? { ...agent, instructions: `${agent.instructions}\n\n${PROMPTS.all}` } : agent;
};

const textOf = (m: ThreadMessage) => m.content.map((p) => (p.type === "text" ? p.text : "")).join("");
type CallPart = { type: "tool-call"; toolName: string; args: Record<string, unknown>; argsText?: string; result?: unknown; isError?: boolean };
const callsOf = (m: ThreadMessage) => m.content.filter((p) => p.type === "tool-call") as unknown as CallPart[];

/** A past message for the planner: its text, the questions it asked (with its plan) and the actions taken. */
function historyOf(m: ThreadMessage) {
  const calls = callsOf(m);
  const asked = calls.find((c) => c.toolName === "ask_user")?.args as Choices | undefined;
  const plan = calls.find((c) => c.toolName === "implement")?.args.plan;
  const actions = calls
    .filter((c) => c.toolName !== "ask_user" && c.toolName !== "implement")
    .map((c) => `  - ${c.toolName}(${(c.argsText ?? "").slice(0, 120)}) → ${c.isError ? "error: " : ""}${String(c.result ?? "").slice(0, 120)}`);
  const files = m.content.filter((p) => p.type === "file").map((p) => (p as { filename?: string }).filename ?? "file");
  return [
    `${m.role}: ${textOf(m)}${files.length ? `[file added: ${files.join(", ")}]` : ""}`,
    asked && `  asked: ${asked.message ? `${asked.message.slice(0, 400)}\n  ` : ""}${asked.questions.map((q) => `${q.question} [${q.options.join(" | ")}]`).join(" / ")}\n  plan proposed: ${asked.plan.slice(0, 600)}`,
    plan && `  plan implemented: ${String(plan).slice(0, 600)}`,
    actions.length && `  actions:\n${actions.join("\n")}`,
  ]
    .filter(Boolean)
    .join("\n");
}

function adapter(agent: { current: Agent }, conversationId: string, scope: string): ChatModelAdapter {
  return {
    async *run({ messages, abortSignal }) {
      const previous = messages.slice(0, -1);
      // The message answers the planner's questions when its last assistant message asked some.
      const lastAssistant = [...previous].reverse().find((m) => m.role === "assistant");
      const answering = !!lastAssistant && callsOf(lastAssistant).some((c) => c.toolName === "ask_user");
      // Files added since the last answer go with this message (also after a reload of the page).
      const since = messages.slice(messages.map((m) => m.role).lastIndexOf("assistant") + 1);
      for (const part of since.flatMap((m) => [...m.content] as { type: string; data?: string }[])) if (part.type === "file") await addFileToConversation(conversationId, part.data ?? "");
      const { model, implementerModel } = await settings.get();
      const startedAt = new Date().toISOString();
      const t0 = performance.now();
      // Full trace kept with the message (exported with the conversation): model exchanges, tool timings.
      const parts = (u: Update, text?: string) => ({
        content: [
          ...u.steps.map((s) => ({ type: "tool-call" as const, toolCallId: s.id, toolName: s.name, args: s.args as never, argsText: JSON.stringify(s.args), result: s.result, isError: s.error })),
          // The private LLM's markdown is not always well formed (one-line lists, fences): fixed here.
          ...(text ? [{ type: "text" as const, text: cleanMarkdown(text) }] : []),
        ],
        metadata: { custom: { status: u.status, trace: { model: model ?? AI_MODELS[0], implementerModel: implementerModel ?? model ?? AI_MODELS[0], startedAt, ms: Math.round(performance.now() - t0), remote: u.remote, steps: u.steps, llm: u.llm } } },
      });
      let latest: Update = { steps: [], llm: [] };
      try {
        for await (const update of runTurn({ agent: withAdmin(agent.current, scope), model: model ?? AI_MODELS[0], implementerModel: implementerModel ?? model ?? AI_MODELS[0], conversationId, history: previous.map(historyOf).join("\n\n"), request: textOf(messages.at(-1)!), answering, signal: abortSignal, scope })) {
          latest = update;
          yield parts(update, update.answer);
        }
        yield parts({ ...latest, status: undefined }, latest.answer);
      } catch (error) {
        if (abortSignal.aborted) throw error;
        // Shown in the conversation (and kept) instead of failing silently.
        yield parts({ ...latest, status: undefined }, `⚠️ ${error instanceof Error ? error.message : String(error)}`);
      }
    },
  };
}

/** The conversation as a JSON file: messages, tool calls (arguments, results, errors, durations) and every model exchange. */
function download(c: Conversation) {
  const json = JSON.stringify({ exportedAt: new Date().toISOString(), ...c, updatedAt: new Date(c.updatedAt).toISOString() }, null, 2);
  const url = URL.createObjectURL(new Blob([json], { type: "application/json" }));
  const name = `${c.scope}-${c.title}`.replace(/[^\p{L}\p{N}-]+/gu, "-").slice(0, 80);
  Object.assign(document.createElement("a"), { href: url, download: `${name}.json` }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** Deletes a conversation: its remote conversations on the LLM API first (when the API can), then the local copy. */
async function deleteConversation(c: Conversation) {
  const remotes = new Set((c.messages as { metadata?: { custom?: { trace?: { remote?: string } } } }[]).map((m) => m.metadata?.custom?.trace?.remote).filter((id): id is string => !!id));
  forgetSession(c.id);
  for (const id of remotes) await deleteRemoteConversation(id).catch((e) => console.warn("Remote conversation not deleted", id, e));
  await remove("conversations", c.id);
}

/** Conversations of this screen: search, open, download, delete. */
function ConversationPicker({ conversations, current, onSelect, onNew, onDelete }: { conversations: Conversation[]; current?: Conversation; onSelect: (id: string) => void; onNew: () => void; onDelete: (c: Conversation) => void }) {
  const [open, setOpen] = useState(false);
  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button variant="ghost" className="h-8 min-w-0 flex-1 justify-between gap-2 px-2 text-[13px] font-normal">
          <span className={cn("truncate", !current && "text-muted-foreground")}>{current?.title ?? "New conversation"}</span>
          <ChevronsUpDownIcon className="size-3.5 shrink-0 text-muted-foreground" />
        </Button>
      </PopoverTrigger>
      <PopoverContent align="start" className="w-[340px] p-0">
        <Command>
          <CommandInput placeholder="Search conversations…" />
          <CommandList>
            <CommandEmpty>No conversation.</CommandEmpty>
            <CommandGroup>
              <CommandItem onSelect={() => (onNew(), setOpen(false))} className="gap-2">
                <PlusIcon className="size-4" /> New conversation
              </CommandItem>
              {conversations.map((c) => (
                <CommandItem key={c.id} value={`${c.title} ${c.id}`} onSelect={() => (onSelect(c.id), setOpen(false))} className="group gap-2">
                  <CheckIcon className={cn("size-4 text-primary", c.id === current?.id ? "opacity-100" : "opacity-0")} />
                  <span className="min-w-0 flex-1 truncate">{c.title}</span>
                  <span className="text-[11px] text-muted-foreground group-hover:hidden">{new Date(c.updatedAt).toLocaleDateString()}</span>
                  <span className="hidden items-center gap-0.5 group-hover:flex">
                    <button type="button" aria-label="Download the conversation (JSON with the model exchanges)" className="icon-btn size-6" onClick={(e) => (e.stopPropagation(), download(c))}>
                      <DownloadIcon />
                    </button>
                    <button type="button" aria-label="Delete the conversation" className="icon-btn size-6 hover:text-destructive" onClick={(e) => (e.stopPropagation(), confirm(`Delete "${c.title}"?`) && onDelete(c))}>
                      <Trash2Icon />
                    </button>
                  </span>
                </CommandItem>
              ))}
            </CommandGroup>
          </CommandList>
        </Command>
      </PopoverContent>
    </Popover>
  );
}

/** `onConversation`: told which conversation is open (the screen can keep its own state with it). */
export function Chat({ scope, agent, placeholder = "Ask anything…", onConversation }: { scope: string; agent: Agent; placeholder?: string; onConversation?: (id: string) => void }) {
  const conversations = useCollection("conversations")
    .filter((c) => c.scope === scope)
    .sort((a, b) => b.updatedAt - a.updatedAt);
  // The last conversation of this screen (report, Excel file…) is reopened.
  const [currentId, setCurrent] = useState<string>(() => last.get(`chat:${scope}`) ?? crypto.randomUUID());
  const setCurrentId = (id: string) => {
    last.set(`chat:${scope}`, id);
    setCurrent(id);
  };
  useEffect(() => setCurrent(last.get(`chat:${scope}`) ?? crypto.randomUUID()), [scope]);
  const current = conversations.find((c) => c.id === currentId);
  useEffect(() => onConversation?.(currentId), [currentId]); // eslint-disable-line react-hooks/exhaustive-deps
  const agentRef = useRef(agent);
  agentRef.current = agent;

  return (
    <div className="flex h-full min-h-0 flex-col">
      <div className="flex h-11 shrink-0 items-center gap-1 border-b px-2">
        <ConversationPicker
          conversations={conversations}
          current={current}
          onSelect={setCurrentId}
          onNew={() => setCurrentId(crypto.randomUUID())}
          onDelete={(c) => {
            void deleteConversation(c);
            if (c.id === currentId) setCurrentId(crypto.randomUUID());
          }}
        />
        <Button variant="ghost" size="icon" className="size-8 text-muted-foreground" aria-label="New conversation" onClick={() => setCurrentId(crypto.randomUUID())}>
          <PlusIcon />
        </Button>
      </div>
      <Runtime key={currentId} id={currentId} scope={scope} initial={current} agent={agentRef} placeholder={placeholder} />
    </div>
  );
}

/** Sends a user message from inside a message (the answers of a question card). */
const Send = createContext<(text: string) => void>(() => undefined);

function Runtime({ id, scope, initial, agent, placeholder }: { id: string; scope: string; initial?: Conversation; agent: { current: Agent }; placeholder: string }) {
  const chatModel = useMemo(() => adapter(agent, id, scope), [agent, id, scope]);
  const runtime = useLocalRuntime(chatModel, { initialMessages: (initial?.messages ?? []) as never });

  // Saves the conversation when a run ends.
  useEffect(
    () =>
      runtime.thread.subscribe(() => {
        const { messages, isRunning } = runtime.thread.getState();
        if (isRunning || messages.length === 0) return;
        const title = textOf(messages.find((m) => m.role === "user") ?? messages[0]).slice(0, 60) || "Conversation";
        const stored = messages.map(({ id, role, content, createdAt, metadata }) => ({
          id,
          role,
          createdAt,
          content: content.filter((p) => p.type === "text" || p.type === "tool-call" || p.type === "file"),
          ...(metadata?.custom && Object.keys(metadata.custom).length ? { metadata: { custom: metadata.custom } } : {}),
        }));
        // What the screen keeps with the conversation (e.g. the viewer's report state) is kept.
        void find("conversations", id).then((old) => put("conversations", { ...old, id, scope, title, messages: stored, updatedAt: Date.now() }));
      }),
    [runtime, id, scope],
  );

  // A file: uploaded to the AI API, added to the conversation (sent with the next message), and shown
  // as a message of its own (no answer is asked for it).
  const fileInput = useRef<HTMLInputElement>(null);
  const [upload, setUpload] = useState<{ busy?: boolean; error?: string }>({});
  const attach = (file: File) => {
    setUpload({ busy: true });
    uploadAiFile(file)
      .then(async (fileId) => {
        await addFileToConversation(id, fileId);
        runtime.thread.append({ role: "user", content: [{ type: "file", filename: file.name, data: fileId, mimeType: file.type || "application/octet-stream", sourceType: "id" }], startRun: false });
        setUpload({});
      })
      .catch((e) => setUpload({ error: e instanceof Error ? e.message : String(e) }));
  };
  const send = (text: string) => runtime.thread.append({ role: "user", content: [{ type: "text", text }] });

  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <Send.Provider value={send}>
        <ThreadPrimitive.Root className="flex min-h-0 flex-1 flex-col">
          <ThreadPrimitive.Viewport className="flex min-h-0 flex-1 flex-col gap-5 overflow-x-hidden overflow-y-auto bg-background px-4 py-5">
            <ThreadPrimitive.Empty>
              <div className="m-auto flex max-w-72 flex-col items-center gap-3 text-center">
                <span className="flex size-10 items-center justify-center rounded-xl bg-primary/10 text-primary">
                  <MessageCircleQuestionIcon className="size-5" />
                </span>
                <p className="text-[13.5px] leading-relaxed text-muted-foreground">{placeholder}</p>
                <p className="text-[12px] leading-relaxed text-muted-foreground/80">The assistant answers questions right away, and asks you to confirm before changing anything.</p>
              </div>
            </ThreadPrimitive.Empty>
            <ThreadPrimitive.Messages components={{ UserMessage, AssistantMessage }} />
          </ThreadPrimitive.Viewport>
          <ComposerPrimitive.Root className="flex flex-col gap-1.5 border-t bg-card p-3">
            {upload.error && <p className="px-2 text-[12px] text-destructive">{upload.error}</p>}
            <div className="flex min-w-0 flex-1 items-end gap-1 rounded-2xl border border-input bg-background py-1 pr-1 pl-1.5 transition-shadow focus-within:border-ring focus-within:ring-[3px] focus-within:ring-ring/20">
              <input
                ref={fileInput}
                type="file"
                hidden
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  e.target.value = "";
                  if (file) attach(file);
                }}
              />
              <Button type="button" variant="ghost" size="icon" aria-label="Add a file to the conversation" disabled={upload.busy} className="size-9 rounded-full text-muted-foreground" onClick={() => fileInput.current?.click()}>
                <PaperclipIcon className={upload.busy ? "animate-pulse" : ""} />
              </Button>
              <ComposerPrimitive.Input placeholder={placeholder} rows={1} className="max-h-40 min-h-9 flex-1 resize-none bg-transparent py-2 text-[14px] outline-none placeholder:text-muted-foreground" />
              <ThreadPrimitive.If running={false}>
                <ComposerPrimitive.Send className="flex size-9 shrink-0 items-center justify-center rounded-full bg-primary text-primary-foreground transition-colors hover:bg-primary/90 disabled:opacity-40" aria-label="Send">
                  <ArrowUpIcon className="size-4" />
                </ComposerPrimitive.Send>
              </ThreadPrimitive.If>
              <ThreadPrimitive.If running>
                <ComposerPrimitive.Cancel className="flex size-9 shrink-0 items-center justify-center rounded-full bg-muted text-foreground hover:bg-accent" aria-label="Stop">
                  <SquareIcon className="size-3.5" />
                </ComposerPrimitive.Cancel>
              </ThreadPrimitive.If>
            </div>
          </ComposerPrimitive.Root>
        </ThreadPrimitive.Root>
      </Send.Provider>
    </AssistantRuntimeProvider>
  );
}

const timeOf = (d?: Date) => (d ? new Date(d).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" }) : "");

/** Your message: a solid bubble on the right, with its time. */
const UserMessage = () => {
  const createdAt = useAuiState((s) => s.message.createdAt);
  const isFile = useAuiState((s) => s.message.content.some((p) => p.type === "file"));
  return (
    <MessagePrimitive.Root className="flex flex-col items-end gap-1">
      {isFile ? (
        <MessagePrimitive.Parts components={{ File: FileChip }} />
      ) : (
        <div className="max-w-[85%] rounded-2xl rounded-br-md bg-primary px-4 py-2.5 text-[14px] leading-relaxed whitespace-pre-wrap text-primary-foreground [overflow-wrap:anywhere] shadow-xs">
          <MessagePrimitive.Parts />
        </div>
      )}
      <span className="pr-1 text-[11px] text-muted-foreground">{timeOf(createdAt)}</span>
    </MessagePrimitive.Root>
  );
};

/** A file added to the conversation: its name, as a message of its own. */
const FileChip = ({ filename, mimeType }: { filename?: string; mimeType: string }) => (
  <span className="flex max-w-[85%] items-center gap-2.5 rounded-2xl rounded-br-md border bg-card px-3.5 py-2.5 text-[13.5px] shadow-xs" title={mimeType}>
    <span className="flex size-8 shrink-0 items-center justify-center rounded-lg bg-muted">
      <FileTextIcon className="size-4 text-muted-foreground" />
    </span>
    <span className="min-w-0 truncate font-medium">{filename ?? "File"}</span>
  </span>
);

type Custom = { status?: Status; trace?: { ms?: number } };

/**
 * The assistant: what it did (folded into one line), its questions (a card), the plan handed to the
 * implementer, what it is doing now, and its answer; under it, its time and how long it took.
 */
const AssistantMessage = () => {
  const running = useAuiState((s) => s.message.status?.type === "running");
  const custom = (useAuiState((s) => s.message.metadata?.custom) ?? {}) as Custom;
  const createdAt = useAuiState((s) => s.message.createdAt);
  const hasText = useAuiState((s) => s.message.content.some((p) => p.type === "text" && p.text));
  const seconds = custom.trace?.ms !== undefined ? (custom.trace.ms / 1000).toFixed(1) : undefined;
  return (
    <MessagePrimitive.Root className="flex flex-col gap-1">
      <div className="flex w-full max-w-full min-w-0 flex-col gap-2.5 rounded-2xl rounded-tl-md border bg-card px-4 py-3 text-[14px] [overflow-wrap:anywhere] shadow-xs [&_pre]:overflow-x-auto [&_pre]:[overflow-wrap:normal] [&>:empty]:hidden">
        <Steps running={running} />
        <MessagePrimitive.Parts components={{ Text: Markdown, tools: { by_name: { ask_user: ChoicesCard, implement: PlanChip }, Fallback: () => null } }} />
        {running && custom.status && <StatusLine status={custom.status} />}
        {running && !custom.status && !hasText && <Dots />}
      </div>
      <span className="pl-1 text-[11px] text-muted-foreground tabular-nums">
        {timeOf(createdAt)}
        {seconds && ` · ${running ? "" : "in "}${seconds} s`}
      </span>
    </MessagePrimitive.Root>
  );
};

/** Nothing for an empty text part (it left a blank gap in the card while the agent works). */
const Markdown = ({ text }: { text: string }) => (text.trim() ? <MarkdownTextPrimitive remarkPlugins={[remarkGfm]} className="markdown" /> : null);

/** The planner's questions: options to click (one question: a click sends it), or a free answer. */
function ChoicesCard({ args }: ToolCallMessagePartProps) {
  const choices = args as unknown as Choices;
  const send = useContext(Send);
  const isLast = useAuiState((s) => s.message.isLast);
  const running = useAuiState((s) => s.message.status?.type === "running");
  const open = isLast && !running;
  const questions = choices.questions ?? [];
  const [picked, setPicked] = useState<string[][]>(() => questions.map(() => []));
  const [free, setFree] = useState<string[]>(() => questions.map(() => ""));
  const single = questions.length === 1 && !questions[0].multiple;
  const answerOf = (i: number) => [...picked[i], ...(free[i].trim() ? [free[i].trim()] : [])];
  const submit = (answers = questions.map((_, i) => answerOf(i))) => send(questions.length === 1 ? answers[0].join(", ") : questions.map((q, i) => `${q.question} → ${answers[i].join(", ") || "(no answer)"}`).join("\n"));
  const toggle = (i: number, o: string) => {
    if (single) return submit([[o]]);
    setPicked((p) => p.map((list, k) => (k !== i ? list : questions[i].multiple ? (list.includes(o) ? list.filter((x) => x !== o) : [...list, o]) : [o])));
  };
  return (
    <div className="flex flex-col gap-3">
      {choices.message && <MarkdownBlock text={choices.message} />}
      <div className={cn("flex flex-col gap-3 rounded-xl border bg-muted/40 p-3", !open && "opacity-70")}>
        {questions.map((q, i) => (
          <div key={i} className="flex flex-col gap-2">
            <span className="flex items-start gap-2 text-[13.5px] font-medium">
              <MessageCircleQuestionIcon className="mt-0.5 size-4 shrink-0 text-primary" aria-hidden />
              {q.question}
            </span>
            <div className="flex flex-wrap gap-1.5 pl-6">
              {q.options.map((o) => {
                const on = picked[i]?.includes(o);
                return (
                  <button
                    key={o}
                    type="button"
                    disabled={!open}
                    aria-pressed={on}
                    onClick={() => toggle(i, o)}
                    className={cn(
                      "rounded-lg border px-3 py-1.5 text-left text-[13px] font-medium transition-colors disabled:cursor-default",
                      on ? "border-primary bg-primary text-primary-foreground" : "border-input bg-card hover:border-primary/50 hover:bg-accent hover:text-accent-foreground",
                    )}
                  >
                    {o}
                  </button>
                );
              })}
            </div>
            {open && (
              <form
                className="flex gap-1.5 pl-6"
                onSubmit={(e) => {
                  e.preventDefault();
                  if (single && free[i].trim()) submit([[free[i].trim()]]);
                }}
              >
                <Input value={free[i]} onChange={(e) => setFree((f) => f.map((v, k) => (k === i ? e.target.value : v)))} placeholder="Other answer…" className="h-8 text-[13px]" aria-label={`Other answer to: ${q.question}`} />
                {single && (
                  <Button type="submit" size="icon" variant="outline" className="size-8 shrink-0" disabled={!free[i].trim()} aria-label="Send this answer">
                    <SendIcon />
                  </Button>
                )}
              </form>
            )}
          </div>
        ))}
        {open && !single && (
          <Button size="sm" className="ml-6 w-fit" disabled={questions.some((_, i) => !answerOf(i).length)} onClick={() => submit()}>
            <SendIcon /> Send my answers
          </Button>
        )}
      </div>
    </div>
  );
}

/** The plan handed to the implementer once the user agreed (unfolded on click). */
function PlanChip({ args }: ToolCallMessagePartProps) {
  const [open, setOpen] = useState(false);
  return (
    <div className="rounded-lg border border-primary/25 bg-primary/5 px-3 py-2">
      <button type="button" onClick={() => setOpen(!open)} className="flex w-full items-center gap-2 text-[12.5px] font-medium text-primary">
        <ListChecksIcon className="size-4" aria-hidden /> Plan validated — implementing
        <ChevronRightIcon className={cn("ml-auto size-3.5 transition-transform", open && "rotate-90")} />
      </button>
      {open && <pre className="mt-2 text-[12px] leading-relaxed whitespace-pre-wrap text-foreground/80">{String((args as { plan?: string }).plan ?? "")}</pre>}
    </div>
  );
}

const MarkdownBlock = ({ text }: { text: string }) => <MarkdownView>{text}</MarkdownView>;

/** "3 actions · 1 failed", unfolded on click: each action, its arguments and result. */
function Steps({ running }: { running: boolean }) {
  const [open, setOpen] = useState(false);
  const [shown, setShown] = useState<number>();
  const content = useAuiState((s) => s.message.content);
  const steps = (content.filter((p) => p.type === "tool-call") as unknown as ToolCallMessagePartProps[]).filter((s) => s.toolName !== "ask_user" && s.toolName !== "implement");
  if (!steps.length) return null;
  const failed = steps.filter((s) => s.isError).length;
  return (
    <div className="border-b pb-2 last:border-b-0 last:pb-0">
      <button type="button" onClick={() => setOpen(!open)} className="flex items-center gap-1.5 text-[12px] text-muted-foreground hover:text-foreground">
        <WrenchIcon className="size-3" aria-hidden />
        {steps.length} action{steps.length > 1 ? "s" : ""}
        {failed > 0 && <span className="text-destructive">· {failed} failed</span>}
        {running && steps.at(-1)?.result === undefined && <span>· {steps.at(-1)?.toolName}…</span>}
        <ChevronRightIcon className={cn("size-3 transition-transform", open && "rotate-90")} />
      </button>
      {open && (
        <ul className="mt-1.5 flex flex-col gap-1">
          {steps.map((s, i) => (
            <li key={i} className="text-[12px]">
              <button type="button" className="flex w-full items-start gap-1.5 text-left" onClick={() => setShown(shown === i ? undefined : i)}>
                {s.result === undefined ? <span className="mt-1 size-2 shrink-0 animate-pulse rounded-full bg-primary" /> : s.isError ? <XIcon className="mt-0.5 size-3 shrink-0 text-destructive" /> : <CheckIcon className="mt-0.5 size-3 shrink-0 text-success" />}
                <span className="min-w-0">
                  <span className="font-mono">{s.toolName}</span>
                  <span className={cn("block truncate", s.isError ? "text-destructive" : "text-muted-foreground")}>{String(s.result ?? "running…").split("\n")[0]}</span>
                </span>
              </button>
              {shown === i && <pre className="mt-1 max-h-60 overflow-auto rounded bg-muted p-1.5 whitespace-pre-wrap">{`${s.argsText}\n\n${String(s.result ?? "…")}`}</pre>}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

/** What the agent is doing now; a bar when its progress is known (e.g. visuals of the page read). */
function StatusLine({ status }: { status: Status }) {
  const known = status.total !== undefined && status.total > 0;
  return (
    <div className="flex min-w-48 flex-col gap-1.5 py-0.5">
      <span className="flex items-center gap-2 text-[13px] text-muted-foreground">
        <span className="size-1.5 animate-pulse rounded-full bg-primary" />
        {status.label}
        {known && <span className="tabular-nums">{status.done ?? 0}/{status.total}</span>}…
      </span>
      {known && (
        <span className="h-1 overflow-hidden rounded-full bg-muted">
          <span className="block h-full rounded-full bg-primary transition-[width] duration-300" style={{ width: `${Math.round(((status.done ?? 0) / status.total!) * 100)}%` }} />
        </span>
      )}
    </div>
  );
}

const Dots = () => (
  <span className="flex gap-1 py-1" aria-label="Thinking">
    {[0, 1, 2].map((i) => (
      <span key={i} className="size-1.5 animate-bounce rounded-full bg-muted-foreground" style={{ animationDelay: `${i * 120}ms` }} />
    ))}
  </span>
);
