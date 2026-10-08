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
import { ArrowUpIcon, ChevronRightIcon, DownloadIcon, PlusIcon, SquareIcon, Trash2Icon } from "lucide-react";
import { runAgent, type Agent, type Update } from "@/agent/loop";
import { AI_MODELS } from "@/agent/llm";
import { put, remove, settings, useCollection, type Conversation } from "@/lib/store";
import { last } from "@/lib/last";

const textOf = (m: ThreadMessage) => m.content.map((p) => (p.type === "text" ? p.text : "")).join("");

function adapter(agent: { current: Agent }): ChatModelAdapter {
  return {
    async *run({ messages, abortSignal }) {
      const previous = messages.slice(0, -1).map((m) => `${m.role}: ${textOf(m)}`).join("\n\n");
      const { model } = await settings.get();
      const startedAt = new Date().toISOString();
      // Full trace kept with the message (exported with the conversation): raw model exchanges, tool timings.
      const parts = (u: Update, text?: string) => ({
        content: [
          ...u.steps.map((s) => ({ type: "tool-call" as const, toolCallId: s.id, toolName: s.name, args: s.args as never, argsText: JSON.stringify(s.args), result: s.result, isError: s.error })),
          ...(text !== undefined ? [{ type: "text" as const, text }] : []),
        ],
        metadata: { custom: { trace: { model: model ?? AI_MODELS[0], startedAt, steps: u.steps, llm: u.llm } } },
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
      <div className="flex items-center gap-1 border-b px-2 py-1.5">
        <select className="h-7 min-w-0 flex-1 truncate rounded-md bg-transparent px-1 text-[12px] outline-none hover:bg-muted" value={current ? currentId : ""} onChange={(e) => setCurrentId(e.target.value || crypto.randomUUID())}>
          <option value="">New conversation</option>
          {conversations.map((c) => (
            <option key={c.id} value={c.id}>
              {c.title}
            </option>
          ))}
        </select>
        <button type="button" title="New conversation" className="icon-btn" onClick={() => setCurrentId(crypto.randomUUID())}>
          <PlusIcon />
        </button>
        {current && (
          <button type="button" title="Download the conversation (JSON with tool calls and model exchanges)" className="icon-btn" onClick={() => download(current)}>
            <DownloadIcon />
          </button>
        )}
        {current && (
          <button type="button" title="Delete conversation" className="icon-btn" onClick={() => void remove("conversations", current.id).then(() => setCurrentId(crypto.randomUUID()))}>
            <Trash2Icon />
          </button>
        )}
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
