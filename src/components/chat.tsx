/** AI chat (assistant-ui) on top of the plain-text agent loop, with conversations kept in IndexedDB. */
import { useEffect, useMemo, useRef, useState } from "react";
import {
  AssistantRuntimeProvider,
  ComposerPrimitive,
  MessagePrimitive,
  ThreadPrimitive,
  useLocalRuntime,
  type ChatModelAdapter,
  type ThreadMessage,
  type ToolCallMessagePartProps,
} from "@assistant-ui/react";
import { MarkdownTextPrimitive } from "@assistant-ui/react-markdown";
import remarkGfm from "remark-gfm";
import { ArrowUpIcon, CheckIcon, ChevronRightIcon, ChevronsUpDownIcon, DownloadIcon, PlusIcon, SearchIcon, SquareIcon, Trash2Icon } from "lucide-react";
import { Popover } from "./popover";
import { runAgent, type Agent, type Update } from "@/agent/loop";
import { AI_MODELS, deleteRemoteConversation } from "@/agent/llm";
import { put, remove, settings, useCollection, type Conversation } from "@/lib/store";
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

function adapter(agent: { current: Agent }): ChatModelAdapter {
  return {
    async *run({ messages, abortSignal }) {
      const previous = messages.slice(0, -1).map(historyOf).join("\n\n");
      const { model } = await settings.get();
      const startedAt = new Date().toISOString();
      // Full trace kept with the message (exported with the conversation): raw model exchanges, tool timings.
      const parts = (u: Update, text?: string) => ({
        content: [
          ...u.steps.map((s) => ({ type: "tool-call" as const, toolCallId: s.id, toolName: s.name, args: s.args as never, argsText: JSON.stringify(s.args), result: s.result, isError: s.error })),
          ...(text !== undefined ? [{ type: "text" as const, text }] : []),
        ],
        metadata: { custom: { trace: { model: model ?? AI_MODELS[0], startedAt, remote: u.remote, steps: u.steps, llm: u.llm } } },
      });
      let last: Update = { steps: [], llm: [] };
      try {
        for await (const update of runAgent({ agent: agent.current, model: model ?? AI_MODELS[0], history: previous, request: textOf(messages.at(-1)!), signal: abortSignal })) {
          last = update;
          yield parts(update, update.answer);
        }
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

export function Chat({ scope, agent, placeholder = "Ask anything…" }: { scope: string; agent: Agent; placeholder?: string }) {
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
  const chatModel = useMemo(() => adapter(agent), [agent]);
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
        void put("conversations", { id, scope, title, messages: stored, updatedAt: Date.now() });
      }),
    [runtime, id, scope],
  );

  return (
    <AssistantRuntimeProvider runtime={runtime}>
      <ThreadPrimitive.Root className="flex min-h-0 flex-1 flex-col">
        <ThreadPrimitive.Viewport className="flex min-h-0 flex-1 flex-col gap-3 overflow-y-auto px-3 py-3">
          <ThreadPrimitive.Empty>
            <p className="m-auto max-w-60 text-center text-[12px] text-muted-foreground">{placeholder}</p>
          </ThreadPrimitive.Empty>
          <ThreadPrimitive.Messages components={{ UserMessage, AssistantMessage }} />
          <ThreadPrimitive.If running>
            <span className="size-2 shrink-0 animate-pulse rounded-full bg-muted-foreground" />
          </ThreadPrimitive.If>
        </ThreadPrimitive.Viewport>
        <ComposerPrimitive.Root className="m-2 flex items-end gap-1 rounded-lg border bg-card p-1.5 focus-within:border-ring">
          <ComposerPrimitive.Input placeholder={placeholder} rows={1} className="max-h-40 min-h-8 flex-1 resize-none bg-transparent px-1.5 py-1 text-[13px] outline-none" />
          <ThreadPrimitive.If running={false}>
            <ComposerPrimitive.Send className="icon-btn bg-primary text-primary-foreground hover:bg-primary/90">
              <ArrowUpIcon />
            </ComposerPrimitive.Send>
          </ThreadPrimitive.If>
          <ThreadPrimitive.If running>
            <ComposerPrimitive.Cancel className="icon-btn">
              <SquareIcon />
            </ComposerPrimitive.Cancel>
          </ThreadPrimitive.If>
        </ComposerPrimitive.Root>
      </ThreadPrimitive.Root>
    </AssistantRuntimeProvider>
  );
}

const UserMessage = () => (
  <MessagePrimitive.Root className="ml-8 self-end rounded-lg bg-muted px-3 py-2 text-[13px] whitespace-pre-wrap">
    <MessagePrimitive.Parts />
  </MessagePrimitive.Root>
);

const AssistantMessage = () => (
  <MessagePrimitive.Root className="flex flex-col gap-1.5 text-[13px]">
    <MessagePrimitive.Parts components={{ Text: Markdown, tools: { Fallback: ToolStep } }} />
  </MessagePrimitive.Root>
);

const Markdown = () => <MarkdownTextPrimitive remarkPlugins={[remarkGfm]} className="markdown" />;

function ToolStep({ toolName, argsText, result, isError }: ToolCallMessagePartProps) {
  const [open, setOpen] = useState(false);
  return (
    <div className="rounded-md border text-[11px]">
      <button type="button" className="flex w-full items-center gap-1 px-2 py-1 text-left text-muted-foreground hover:text-foreground" onClick={() => setOpen(!open)}>
        <ChevronRightIcon className={`size-3 transition-transform ${open ? "rotate-90" : ""}`} />
        <span className={isError ? "text-destructive" : ""}>{toolName}</span>
        <span className="truncate opacity-70">{argsText}</span>
        {result === undefined && <span className="ml-auto size-1.5 animate-pulse rounded-full bg-muted-foreground" />}
      </button>
      {open && <pre className="max-h-60 overflow-auto border-t px-2 py-1 whitespace-pre-wrap">{`${argsText}\n\n${String(result ?? "…")}`}</pre>}
    </div>
  );
}
