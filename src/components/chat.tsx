/** AI chat (assistant-ui) on top of the plain-text agent loop, with conversations kept in IndexedDB. */
import { useEffect, useMemo, useRef, useState } from "react";
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
import { ArrowUpIcon, CheckIcon, ChevronRightIcon, ChevronsUpDownIcon, DownloadIcon, PlusIcon, SearchIcon, SparklesIcon, SquareIcon, Trash2Icon, WrenchIcon, XIcon } from "lucide-react";
import { Popover } from "./popover";
import { forgetSession, runAgent, type Agent, type Status, type Update } from "@/agent/loop";
import { PROMPTS, withPrompts, type PromptScreen } from "@/config";

/** The screen's agent with the instructions the admin added in Settings (for all screens, and this one). */
const withAdmin = (agent: Agent, scope: string): Agent => {
  const screen = SCREEN_OF[scope.split(":")[0]];
  return screen ? { ...agent, instructions: withPrompts(screen, agent.instructions) } : PROMPTS.all.trim() ? { ...agent, instructions: `${agent.instructions}\n\n${PROMPTS.all}` } : agent;
};
import { AI_MODELS, deleteRemoteConversation } from "@/agent/llm";
import { find, put, remove, settings, useCollection, type Conversation } from "@/lib/store";
import { last } from "@/lib/last";

const textOf = (m: ThreadMessage) => m.content.map((p) => (p.type === "text" ? p.text : "")).join("");

/** A past message for the model: its text, and for the assistant a short list of the actions it took. */
function historyOf(m: ThreadMessage) {
  const actions = m.content
    .filter((p) => p.type === "tool-call")
    .map((p) => {
      const call = p as { toolName: string; argsText?: string; result?: unknown; isError?: boolean };
      return `  - ${call.toolName}(${(call.argsText ?? "").slice(0, 120)}) → ${call.isError ? "error: " : ""}${String(call.result ?? "").slice(0, 120)}`;
    });
  return `${m.role}: ${textOf(m)}${actions.length ? `\n  actions:\n${actions.join("\n")}` : ""}`;
}

/** Screen of a chat scope ("viewer:<id>", "dataset:<id>"…), for the admin's additional instructions. */
const SCREEN_OF: Record<string, PromptScreen> = { viewer: "viewer", builder: "builder", dataset: "datasets", dax: "dax", pivot: "pivot" };

function adapter(agent: { current: Agent }, conversationId: string, scope: string): ChatModelAdapter {
  return {
    async *run({ messages, abortSignal }) {
      const previous = messages.slice(0, -1).map(historyOf).join("\n\n");
      const { model } = await settings.get();
      const startedAt = new Date().toISOString();
      const t0 = performance.now();
      // Full trace kept with the message (exported with the conversation): raw model exchanges, tool timings,
      // and the whole time the answer took (reading the page, the model, the actions). `status`: what is
      // happening now (shown while running, gone in the last update).
      const parts = (u: Update, text?: string) => ({
        content: [
          ...u.steps.map((s) => ({ type: "tool-call" as const, toolCallId: s.id, toolName: s.name, args: s.args as never, argsText: JSON.stringify(s.args), result: s.result, isError: s.error })),
          ...(text !== undefined ? [{ type: "text" as const, text }] : []),
        ],
        metadata: { custom: { status: u.status, trace: { model: model ?? AI_MODELS[0], startedAt, ms: Math.round(performance.now() - t0), remote: u.remote, steps: u.steps, llm: u.llm } } },
      });
      let last: Update = { steps: [], llm: [] };
      try {
        for await (const update of runAgent({ agent: withAdmin(agent.current, scope), model: model ?? AI_MODELS[0], conversationId, history: previous, request: textOf(messages.at(-1)!), signal: abortSignal })) {
          last = update;
          yield parts(update, update.answer);
        }
        // The final time (the last update may have been sent before the end of the run).
        yield parts({ ...last, status: undefined }, last.answer);
      } catch (error) {
        if (abortSignal.aborted) throw error;
        // Shown in the conversation (and kept) instead of failing silently.
        yield parts(last, `⚠️ ${error instanceof Error ? error.message : String(error)}`);
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
  const remotes = new Set(
    (c.messages as { metadata?: { custom?: { trace?: { remote?: string } } } }[]).map((m) => m.metadata?.custom?.trace?.remote).filter((id): id is string => !!id),
  );
  forgetSession(c.id);
  for (const id of remotes) await deleteRemoteConversation(id).catch((e) => console.warn("Remote conversation not deleted", id, e));
  await remove("conversations", c.id);
}

/** Conversations of this screen (shadcn-style combobox): search, open, download, delete. */
function ConversationPicker({ conversations, current, onSelect, onNew, onDelete }: {
  conversations: Conversation[];
  current?: Conversation;
  onSelect: (id: string) => void;
  onNew: () => void;
  onDelete: (c: Conversation) => void;
}) {
  const [search, setSearch] = useState("");
  const shown = conversations.filter((c) => c.title.toLowerCase().includes(search.toLowerCase()));
  return (
    <span className="min-w-0 flex-1">
      <Popover
        full
        width={340}
        trigger={(open) => (
          <button type="button" onClick={() => (setSearch(""), open())} className="flex h-7 w-full items-center gap-1.5 rounded-md px-1.5 text-left text-[12px] hover:bg-muted">
            <span className={`min-w-0 flex-1 truncate ${current ? "" : "text-muted-foreground"}`}>{current?.title ?? "New conversation"}</span>
            <ChevronsUpDownIcon className="size-3.5 shrink-0 text-muted-foreground" />
          </button>
        )}
      >
        {(close) => (
          <>
            <label className="flex items-center gap-2 border-b px-1 pb-1.5">
              <SearchIcon className="size-3.5 text-muted-foreground" />
              <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search conversations…" className="h-7 flex-1 bg-transparent outline-none" />
            </label>
            <div className="max-h-80 overflow-y-auto">
              <button type="button" onClick={() => (onNew(), close())} className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left hover:bg-muted">
                <PlusIcon className="size-3.5" /> New conversation
              </button>
              {shown.map((c) => (
                <div key={c.id} className="group flex items-center gap-1 rounded-md pr-1 hover:bg-muted">
                  <button type="button" onClick={() => (onSelect(c.id), close())} className="flex min-w-0 flex-1 items-center gap-2 px-2 py-1.5 text-left">
                    <CheckIcon className={`size-3.5 shrink-0 ${c.id === current?.id ? "" : "invisible"}`} />
                    <span className="min-w-0 flex-1 truncate">{c.title}</span>
                    <span className="shrink-0 text-[11px] text-muted-foreground">{new Date(c.updatedAt).toLocaleDateString()}</span>
                  </button>
                  <button type="button" title="Download (JSON with tool calls and model exchanges)" className="icon-btn size-6" onClick={() => download(c)}>
                    <DownloadIcon />
                  </button>
                  <button type="button" title="Delete" className="icon-btn size-6 hover:text-destructive" onClick={() => confirm(`Delete "${c.title}"?`) && onDelete(c)}>
                    <Trash2Icon />
                  </button>
                </div>
              ))}
              {shown.length === 0 && <p className="px-2 py-2 text-[12px] text-muted-foreground">No conversation.</p>}
            </div>
          </>
        )}
      </Popover>
    </span>
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
      <div className="flex h-10 shrink-0 items-center gap-1 border-b px-2">
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
        <button type="button" title="New conversation" className="icon-btn" onClick={() => setCurrentId(crypto.randomUUID())}>
          <PlusIcon />
        </button>
      </div>
      <Runtime key={currentId} id={currentId} scope={scope} initial={current} agent={agentRef} placeholder={placeholder} />
    </div>
  );
}

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
          content: content.filter((p) => p.type === "text" || p.type === "tool-call"),
          ...(metadata?.custom && Object.keys(metadata.custom).length ? { metadata: { custom: metadata.custom } } : {}),
        }));
        // What the screen keeps with the conversation (e.g. the viewer's report state) is kept.
        void find("conversations", id).then((old) => put("conversations", { ...old, id, scope, title, messages: stored, updatedAt: Date.now() }));
      }),
    [runtime, id, scope],
  );

  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <ThreadPrimitive.Root className="flex min-h-0 flex-1 flex-col">
        <ThreadPrimitive.Viewport className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto bg-zinc-50 px-3 py-4">
          <ThreadPrimitive.Empty>
            <p className="m-auto max-w-60 text-center text-[12px] text-muted-foreground">{placeholder}</p>
          </ThreadPrimitive.Empty>
          <ThreadPrimitive.Messages components={{ UserMessage, AssistantMessage }} />
        </ThreadPrimitive.Viewport>
        <ComposerPrimitive.Root className="flex items-end gap-1 border-t bg-card p-2.5">
          <div className="flex min-w-0 flex-1 items-end gap-1 rounded-[20px] border bg-zinc-50 py-1 pr-1 pl-4 focus-within:border-ring">
            <ComposerPrimitive.Input placeholder={placeholder} rows={1} className="max-h-40 min-h-8 flex-1 resize-none bg-transparent py-1.5 text-[13px] outline-none" />
            <ThreadPrimitive.If running={false}>
              <ComposerPrimitive.Send className="flex size-8 shrink-0 items-center justify-center rounded-full bg-primary text-primary-foreground hover:bg-primary/90 disabled:opacity-40" title="Send">
                <ArrowUpIcon className="size-4" />
              </ComposerPrimitive.Send>
            </ThreadPrimitive.If>
            <ThreadPrimitive.If running>
              <ComposerPrimitive.Cancel className="flex size-8 shrink-0 items-center justify-center rounded-full bg-muted text-foreground hover:bg-zinc-200" title="Stop">
                <SquareIcon className="size-3.5" />
              </ComposerPrimitive.Cancel>
            </ThreadPrimitive.If>
          </div>
        </ComposerPrimitive.Root>
      </ThreadPrimitive.Root>
    </AssistantRuntimeProvider>
  );
}

const timeOf = (d?: Date) => (d ? new Date(d).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" }) : "");

/** Your message: a solid bubble on the right, with its time. */
const UserMessage = () => {
  const createdAt = useAuiState((s) => s.message.createdAt);
  return (
    <MessagePrimitive.Root className="flex flex-col items-end gap-1">
      <div className="max-w-[85%] rounded-2xl rounded-br-md bg-primary px-3.5 py-2 text-[13px] leading-relaxed whitespace-pre-wrap text-primary-foreground shadow-sm">
        <MessagePrimitive.Parts />
      </div>
      <span className="pr-1 text-[10px] text-muted-foreground">{timeOf(createdAt)}</span>
    </MessagePrimitive.Root>
  );
};

type Custom = { status?: Status; trace?: { ms?: number } };

/**
 * The assistant: an avatar and a bordered bubble holding what it did (folded into one line), what it
 * is doing now (with a progress bar when known), and its answer; under it, its time and how long the
 * whole answer took (reading the page, the model, the actions).
 */
const AssistantMessage = () => {
  const running = useAuiState((s) => s.message.status?.type === "running");
  // Raw values selected (a new object or array per call would draw the message again and again).
  const custom = (useAuiState((s) => s.message.metadata?.custom) ?? {}) as Custom;
  const createdAt = useAuiState((s) => s.message.createdAt);
  const hasText = useAuiState((s) => s.message.content.some((p) => p.type === "text" && p.text));
  const seconds = custom.trace?.ms !== undefined ? (custom.trace.ms / 1000).toFixed(1) : undefined;
  return (
    <MessagePrimitive.Root className="flex items-start gap-2">
      <span className="mt-0.5 flex size-7 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-violet-500 to-blue-500 text-white">
        <SparklesIcon className="size-3.5" />
      </span>
      <div className="flex min-w-0 flex-1 flex-col gap-1">
        <div className="w-fit max-w-full rounded-2xl rounded-tl-md border bg-white px-3.5 py-2.5 text-[13px] shadow-sm">
          <Steps running={running} />
          {running && custom.status && <StatusLine status={custom.status} />}
          {running && !custom.status && !hasText && <Dots />}
          <MessagePrimitive.Parts components={{ Text: Markdown, tools: { Fallback: () => null } }} />
        </div>
        <span className="pl-1 text-[10px] text-muted-foreground tabular-nums">
          {timeOf(createdAt)}
          {seconds && ` · ${running ? "" : "answered in "}${seconds} s`}
        </span>
      </div>
    </MessagePrimitive.Root>
  );
};

const Markdown = () => <MarkdownTextPrimitive remarkPlugins={[remarkGfm]} className="markdown" />;

/** "3 actions · 1 failed", unfolded on click: each action, its arguments and result. */
function Steps({ running }: { running: boolean }) {
  const [open, setOpen] = useState(false);
  const [shown, setShown] = useState<number>();
  const content = useAuiState((s) => s.message.content);
  const steps = content.filter((p) => p.type === "tool-call") as unknown as ToolCallMessagePartProps[];
  if (!steps.length) return null;
  const failed = steps.filter((s) => s.isError).length;
  return (
    <div className="mb-2 border-b pb-2">
      <button type="button" onClick={() => setOpen(!open)} className="flex items-center gap-1.5 text-[11px] text-muted-foreground hover:text-foreground">
        <WrenchIcon className="size-3" />
        {steps.length} action{steps.length > 1 ? "s" : ""}
        {failed > 0 && <span className="text-red-600">· {failed} failed</span>}
        {running && steps.at(-1)?.result === undefined && <span>· {steps.at(-1)?.toolName}…</span>}
        <ChevronRightIcon className={`size-3 transition-transform ${open ? "rotate-90" : ""}`} />
      </button>
      {open && (
        <ul className="mt-1.5 flex flex-col gap-1">
          {steps.map((s, i) => (
            <li key={i} className="text-[11px]">
              <button type="button" className="flex w-full items-start gap-1.5 text-left" onClick={() => setShown(shown === i ? undefined : i)}>
                {s.result === undefined ? <span className="mt-1 size-2 shrink-0 animate-pulse rounded-full bg-blue-500" /> : s.isError ? <XIcon className="mt-0.5 size-3 shrink-0 text-red-500" /> : <CheckIcon className="mt-0.5 size-3 shrink-0 text-emerald-600" />}
                <span className="min-w-0">
                  <span className="font-mono">{s.toolName}</span>
                  <span className={`block truncate ${s.isError ? "text-red-600" : "text-muted-foreground"}`}>{String(s.result ?? "running…").split("\n")[0]}</span>
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
      <span className="flex items-center gap-2 text-[12px] text-muted-foreground">
        <span className="size-1.5 animate-pulse rounded-full bg-blue-500" />
        {status.label}
        {known && <span className="tabular-nums">{status.done ?? 0}/{status.total}</span>}…
      </span>
      {known && (
        <span className="h-1 overflow-hidden rounded-full bg-muted">
          <span className="block h-full rounded-full bg-blue-500 transition-[width] duration-300" style={{ width: `${Math.round(((status.done ?? 0) / status.total!) * 100)}%` }} />
        </span>
      )}
    </div>
  );
}

const Dots = () => (
  <span className="flex gap-1 py-1">
    {[0, 1, 2].map((i) => (
      <span key={i} className="size-1.5 animate-bounce rounded-full bg-muted-foreground" style={{ animationDelay: `${i * 120}ms` }} />
    ))}
  </span>
);
