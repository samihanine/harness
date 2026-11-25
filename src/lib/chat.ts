/**
 * A screen's chat with its agent (the server runs it): the message goes to POST /api/chat with what the
 * screen shows (scope), the answer streams back as events — text, tool steps, questions with buttons,
 * effects on the screen (a new document, a report to reload) and browser actions the server asks for
 * (the embedded report is here), answered at once. Conversations are kept in this browser; the open one is
 * in the address (?c=…), so each tab has its own. Each conversation runs on its own: another can be started
 * or opened while one is answering. A conversation keeps the report's view (page, filters) of its last
 * answer and shows it again when it is opened.
 */
import { useCallback, useEffect, useRef, useSyncExternalStore } from "react";
import type { ChatEvent, ChatItem, ChatScope, FileRef, ScreenId } from "@shared/types";
import { download, errorText, post } from "./api";
import { describeError } from "@/tools/browser/report-actions";
import { ensureFiles, item, projectsOf, readLocal, remove, save, useItems, writeLocal, type Stored } from "./store";
import { modelsFor } from "./items";
import { loadSettings } from "./settings";
import { useScreenSearch } from "./params";

/** Without a choice yet: the default project (the settings' Confluence folder), when there is one. */
const defaultProjects = async () => (await loadSettings().catch(() => undefined))?.confluence.projects.slice(0, 1).map((p) => p.id) ?? [];

export type BrowserActions = Record<string, (args: any) => Promise<unknown>>;
type Conversation = Stored["conversation"];
type Assistant = Extract<ChatItem, { kind: "assistant" }>;

/* ------------------------- Runs (one per conversation) ------------------------- */

type Run = { abort: AbortController; status?: string };
const runs = new Map<string, Run>();
const runListeners = new Set<() => void>();
let runsVersion = 0;
const changed = () => (runsVersion++, runListeners.forEach((l) => l()));
const useRuns = () => useSyncExternalStore((l) => (runListeners.add(l), () => void runListeners.delete(l)), () => runsVersion);

const transcript = (items: ChatItem[]) =>
  items
    .map((i) => (i.kind === "user" ? `User: ${i.text}` : `Assistant: ${i.text}${i.ask ? `\n(asked: ${i.ask.questions.map((q) => q.question).join(" / ")})` : ""}`))
    .join("\n\n")
    .slice(-20_000);

export type ChatOptions = {
  screen: ScreenId;
  item: string;
  scope: () => ChatScope | Promise<ChatScope>;
  actions?: () => BrowserActions;
  onEffect?: (name: string, data: unknown) => void;
  /** The report's view now (a bookmark state), kept with the conversation after each answer. */
  view?: () => Promise<string | undefined>;
  /** Shows a kept view again (the conversation was opened). */
  restoreView?: (state: string) => Promise<void>;
};

export function useChat(o: ChatOptions) {
  useRuns();
  const all = useItems("conversation");
  const [search, setSearch] = useScreenSearch();
  const lastKey = (item: string) => `chat:${o.screen}:${item}`;
  const currentId = search.c ?? readLocal(lastKey(o.item), "");
  const opts = useRef(o);
  opts.current = o;

  // Every conversation of the assistant (all items), the latest first.
  const conversations = (all ?? []).filter((c) => c.screen === o.screen).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  const conversation = conversations.find((c) => c.id === currentId && c.item === o.item);

  /** Opens a conversation (of another item: that item is shown too), or a new one (""). */
  const select = useCallback(
    (id: string, itemId?: string) => {
      writeLocal(lastKey(itemId ?? opts.current.item), id);
      setSearch({ c: id || undefined, ...(itemId && itemId !== opts.current.item ? { item: itemId } : {}) });
    },
    [setSearch], // eslint-disable-line react-hooks/exhaustive-deps
  );
  const newConversation = useCallback(() => select(""), [select]);

  // The view of an opened conversation shown again (once per opening).
  const restored = useRef<string>(undefined);
  useEffect(() => {
    if (!conversation?.view || restored.current === conversation.id || runs.has(conversation.id)) return;
    restored.current = conversation.id;
    void opts.current.restoreView?.(conversation.view).catch(() => undefined);
  }, [conversation?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const send = useCallback(
    async (text: string, files: FileRef[] = []) => {
      const { screen, item: itemId } = opts.current;
      if ((currentId && runs.has(currentId)) || (!text.trim() && !files.length)) return;
      const existing = currentId ? await item("conversation", currentId) : undefined;
      const now = new Date().toISOString();
      let conv: Conversation = existing && existing.item === itemId ? existing : { id: crypto.randomUUID(), screen, item: itemId, title: text.slice(0, 60) || files[0]?.name || "Conversation", items: [], updatedAt: now, createdAt: now };
      const history = transcript(conv.items);
      const answer: Assistant = { kind: "assistant", text: "", tools: [], running: true };
      conv = { ...conv, items: [...conv.items, { kind: "user", text, files: files.length ? files : undefined }, answer], updatedAt: now };
      const id = conv.id;
      restored.current = id;
      select(id);
      await save("conversation", conv);
      const run: Run = { abort: new AbortController(), status: "Thinking…" };
      runs.set(id, run);
      changed();
      const setStatus = (s?: string) => ((run.status = s), changed());
      const ctrl = run.abort;
      // A note written after tool steps starts a new paragraph.
      let afterTool = false;
      // Saved as it streams (a few times a second at most), over the latest copy (rating, view…).
      let last = 0;
      const update = async (force = false) => {
        if (!force && Date.now() - last < 250) return;
        last = Date.now();
        const kept = (await item("conversation", id)) ?? conv;
        conv = { ...kept, items: [...conv.items.slice(0, -1), { ...answer, tools: [...answer.tools] }], updatedAt: new Date().toISOString() };
        await save("conversation", conv);
      };
      try {
        const base = await opts.current.scope();
        // With the projects chosen in the chat header, and the models already read (not read again).
        const projects = projectsOf() ?? (await defaultProjects());
        const ids = [base.report?.datasetId, base.dataset?.id, ...(base.datasets ?? []).map((d) => d.id)];
        const scope: ChatScope = { ...base, projects, models: { ...(await modelsFor(ids)), ...(base.models ?? {}) } };
        await ensureFiles([...files, ...(scope.files ?? [])]);
        const r = await fetch("/api/chat", { method: "POST", signal: ctrl.signal, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ conversationId: id, screen, message: text, files, scope, transcript: history || undefined }) });
        if (!r.ok || !r.body) throw new Error((await r.text()) || `Error ${r.status}`);
        const reader = r.body.getReader();
        const dec = new TextDecoder();
        let buf = "";
        for (;;) {
          // The server sends a ping every 15 s: nothing for 45 s means the connection was lost (server restarted…).
          let timer: ReturnType<typeof setTimeout> | undefined;
          const silent = new Promise<never>((_, reject) => (timer = setTimeout(() => reject(new Error("The connection to the server was lost: send the message again.")), 45_000)));
          const { done, value } = await Promise.race([reader.read(), silent]).finally(() => clearTimeout(timer));
          if (done) break;
          buf += dec.decode(value, { stream: true });
          let at: number;
          while ((at = buf.indexOf("\n\n")) >= 0) {
            const chunk = buf.slice(0, at);
            buf = buf.slice(at + 2);
            const line = chunk.split("\n").find((l) => l.startsWith("data: "));
            if (!line) continue;
            await handle(JSON.parse(line.slice(6)) as ChatEvent);
          }
        }
      } catch (e) {
        if (!ctrl.signal.aborted) answer.error = errorText(e);
        ctrl.abort();
      } finally {
        answer.running = false;
        for (const t of answer.tools) if (t.status === "running") t.status = "error";
        await update(true);
        // The report's view kept with the conversation (shown again when it is opened).
        if (opts.current.item === itemId && opts.current.view) {
          const view = await opts.current.view().catch(() => undefined);
          if (view) await save("conversation", { ...((await item("conversation", id)) ?? conv), view });
        }
        runs.delete(id);
        changed();
      }

      async function handle(e: ChatEvent) {
        switch (e.type) {
          case "text":
            if (afterTool && answer.text && !answer.text.endsWith("\n")) answer.text += "\n\n";
            afterTool = false;
            answer.text += e.text;
            setStatus(undefined);
            return update();
          case "status":
            setStatus(e.label);
            return;
          case "tool": {
            const i = answer.tools.findIndex((t) => t.id === e.id);
            const step = { id: e.id, name: e.name, args: e.args, status: e.status, result: e.result, images: e.images, ms: e.ms };
            i >= 0 ? (answer.tools[i] = step) : answer.tools.push(step);
            afterTool = true;
            setStatus(e.status === "running" ? toolLabel(e.name) : "Thinking…");
            return update();
          }
          case "client_tool": {
            // The screen may now show another item (the user switched): its report is not this conversation's.
            const here = opts.current.item === itemId;
            const action = here ? opts.current.actions?.()[e.name] : undefined;
            const reply = action ? await action(e.args).then((result) => ({ result: result ?? null }), (err) => ({ error: describeError(err) })) : { error: here ? `This screen cannot do "${e.name}" (open the report first).` : "The user opened another item on this screen: this conversation's report is not shown any more. Say so and stop." };
            await post("/chat/tool-result", { callId: e.callId, ...reply }).catch(() => undefined);
            return;
          }
          case "effect":
            if (e.name === "ask") answer.ask = e.data as Assistant["ask"];
            else if (opts.current.item === itemId) opts.current.onEffect?.(e.name, e.data);
            return update(true);
          case "error":
            answer.error = e.error;
            return update(true);
          case "done":
          case "start":
            return;
        }
      }
    },
    [currentId, select],
  );

  const stop = useCallback(() => runs.get(currentId)?.abort.abort(), [currentId]);
  const deleteConversation = useCallback(
    async (id: string) => {
      runs.get(id)?.abort.abort();
      await remove("conversation", id);
      await fetch(`/api/chat/${id}`, { method: "DELETE" }).catch(() => undefined);
      if (id === currentId) newConversation();
    },
    [currentId, newConversation],
  );
  /** The user's feedback on the open conversation. */
  const rate = useCallback(async (patch: Pick<Conversation, "rating" | "comment">) => {
    if (!conversation) return;
    await save("conversation", { ...((await item("conversation", conversation.id)) ?? conversation), ...patch });
  }, [conversation]);
  /** The open conversation as a JSON file: messages, steps, rating and comment. */
  const exportJson = useCallback(() => {
    if (!conversation) return;
    const name = `${conversation.screen}-${conversation.title.replace(/[^\p{L}\p{N}]+/gu, "-").slice(0, 40)}-${conversation.updatedAt.slice(0, 10)}.json`;
    const { view: _view, ...rest } = conversation;
    download(new Blob([JSON.stringify({ ...rest, exportedAt: new Date().toISOString() }, null, 2)], { type: "application/json" }), name);
  }, [conversation]);

  const running = !!currentId && runs.has(currentId);
  return {
    conversation,
    conversations,
    send,
    stop,
    running,
    status: running ? runs.get(currentId)?.status : undefined,
    /** Conversations answering now (also in the background). */
    isRunning: (id: string) => runs.has(id),
    select,
    newConversation,
    deleteConversation,
    rate,
    exportJson,
  };
}

/** What a running tool is doing, in words. */
export const toolLabel = (name: string) =>
  ({
    look_at_report: "Looking at the report…",
    look_at_deck: "Drawing the slides…",
    look_at_slides: "Drawing the slides…",
    run_dax: "Querying the model…",
    model_structure: "Reading the model…",
    save_report: "Saving the report…",
    capture_powerbi: "Taking a picture of the report…",
    sql: "Running SQL…",
    load_dax: "Loading data…",
    read_view: "Reading the report…",
  })[name] ?? "Working…";
