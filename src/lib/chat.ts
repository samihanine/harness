/**
 * A screen's chat with its agent (the server runs it): the message goes to POST /api/chat with what the
 * screen shows (scope), the answer streams back as events — text, tool steps, questions with buttons,
 * effects on the screen (a new document, a report to reload) and browser actions the server asks for
 * (the embedded report is here), answered at once. Conversations are kept in this browser, per item.
 */
import { useCallback, useEffect, useRef, useState } from "react";
import type { ChatEvent, ChatItem, ChatScope, FileRef, ScreenId } from "@shared/types";
import { errorText, post } from "./api";
import { ensureFiles, item, readLocal, remove, save, useItems, writeLocal, type Stored } from "./store";

export type BrowserActions = Record<string, (args: any) => Promise<unknown>>;
type Conversation = Stored["conversation"];
type Assistant = Extract<ChatItem, { kind: "assistant" }>;

const transcript = (items: ChatItem[]) =>
  items
    .map((i) => (i.kind === "user" ? `User: ${i.text}` : `Assistant: ${i.text}${i.ask ? `\n(asked: ${i.ask.questions.map((q) => q.question).join(" / ")})` : ""}`))
    .join("\n\n")
    .slice(-20_000);

export function useChat(o: { screen: ScreenId; item: string; scope: () => ChatScope; actions?: () => BrowserActions; onEffect?: (name: string, data: unknown) => void }) {
  const all = useItems("conversation");
  const key = `chat:${o.screen}:${o.item}`;
  const [currentId, setCurrentId] = useState<string>(() => readLocal(key, ""));
  const [running, setRunning] = useState(false);
  const [status, setStatus] = useState<string>();
  const abort = useRef<AbortController | undefined>(undefined);
  const opts = useRef(o);
  opts.current = o;

  useEffect(() => setCurrentId(readLocal(key, "")), [key]);
  const conversations = (all ?? []).filter((c) => c.screen === o.screen && c.item === o.item).sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  const conversation = conversations.find((c) => c.id === currentId);

  const select = useCallback((id: string) => (setCurrentId(id), writeLocal(key, id)), [key]);
  const newConversation = useCallback(() => select(""), [select]);

  const send = useCallback(
    async (text: string, files: FileRef[] = []) => {
      if (running || (!text.trim() && !files.length)) return;
      const { screen, item: itemId } = opts.current;
      let conv: Conversation = (currentId && (await item("conversation", currentId))) || { id: crypto.randomUUID(), screen, item: itemId, title: text.slice(0, 60) || files[0]?.name || "Conversation", items: [], updatedAt: new Date().toISOString() };
      const history = transcript(conv.items);
      const answer: Assistant = { kind: "assistant", text: "", tools: [], running: true };
      conv = { ...conv, items: [...conv.items, { kind: "user", text, files: files.length ? files : undefined }, answer], updatedAt: new Date().toISOString() };
      select(conv.id);
      await save("conversation", conv);
      setRunning(true);
      setStatus("Thinking…");
      const ctrl = new AbortController();
      abort.current = ctrl;
      // Saved as it streams (a few times a second at most).
      let last = 0;
      const update = async (force = false) => {
        conv = { ...conv, items: [...conv.items.slice(0, -1), { ...answer, tools: [...answer.tools] }], updatedAt: new Date().toISOString() };
        if (force || Date.now() - last > 250) {
          last = Date.now();
          await save("conversation", conv);
        }
      };
      try {
        const scope = opts.current.scope();
        await ensureFiles([...files, ...(scope.files ?? [])]);
        const r = await fetch("/api/chat", { method: "POST", signal: ctrl.signal, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ conversationId: conv.id, screen, message: text, files, scope, transcript: history || undefined }) });
        if (!r.ok || !r.body) throw new Error((await r.text()) || `Error ${r.status}`);
        const reader = r.body.getReader();
        const dec = new TextDecoder();
        let buf = "";
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          buf += dec.decode(value, { stream: true });
          let at: number;
          while ((at = buf.indexOf("\n\n")) >= 0) {
            const chunk = buf.slice(0, at);
            buf = buf.slice(at + 2);
            const line = chunk.split("\n").find((l) => l.startsWith("data: "));
            if (!line) continue;
            const e = JSON.parse(line.slice(6)) as ChatEvent;
            await handle(e);
          }
        }
      } catch (e) {
        if (!ctrl.signal.aborted) answer.error = errorText(e);
      } finally {
        answer.running = false;
        for (const t of answer.tools) if (t.status === "running") t.status = "error";
        await update(true);
        setRunning(false);
        setStatus(undefined);
      }

      async function handle(e: ChatEvent) {
        switch (e.type) {
          case "text":
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
            setStatus(e.status === "running" ? toolLabel(e.name) : "Thinking…");
            return update();
          }
          case "client_tool": {
            const action = opts.current.actions?.()[e.name];
            const reply = action ? await action(e.args).then((result) => ({ result: result ?? null }), (err) => ({ error: errorText(err) })) : { error: `This screen cannot do "${e.name}" (open the report first).` };
            await post("/chat/tool-result", { callId: e.callId, ...reply }).catch(() => undefined);
            return;
          }
          case "effect":
            if (e.name === "ask") answer.ask = e.data as Assistant["ask"];
            else opts.current.onEffect?.(e.name, e.data);
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
    [running, currentId, select],
  );

  const stop = useCallback(() => abort.current?.abort(), []);
  const deleteConversation = useCallback(async (id: string) => {
    await remove("conversation", id);
    await fetch(`/api/chat/${id}`, { method: "DELETE" }).catch(() => undefined);
    if (id === currentId) newConversation();
  }, [currentId, newConversation]);

  return { conversation, conversations, send, stop, running, status, select, newConversation, deleteConversation };
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
