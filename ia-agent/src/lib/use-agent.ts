/** State of the agent panel: host connection, conversations, runs. */
import { useCallback, useEffect, useRef, useState } from "react";
import { Host } from "@/harness/host";
import type { HostPrompt } from "@/harness/host";
import { historyText } from "@/harness/history";
import { message, newStep, runAgent } from "@/harness/run";
import { AI_MODELS } from "./llm";
import { conversations, settings } from "./store";
import type { AgentItem, Conversation, Step } from "./types";

export type HostState =
  | { status: "connecting" }
  | { status: "standalone" }
  | { status: "connected"; host: Host; title: string; tools: number; resources: number; prompts: HostPrompt[] };

const blank = (scope: string): Conversation => ({
  id: crypto.randomUUID(),
  scope,
  title: "New conversation",
  createdAt: Date.now(),
  updatedAt: Date.now(),
  items: [],
});

export function useAgent() {
  const [host, setHost] = useState<HostState>({ status: "connecting" });
  const scope = host.status === "connected" ? host.host.info.name : "standalone";
  const [conversation, setConversation] = useState<Conversation>(() => blank(scope));
  const [list, setList] = useState<Omit<Conversation, "items">[]>([]);
  const [model, setModelState] = useState(AI_MODELS[0]);
  const abort = useRef<AbortController | null>(null);
  const approvals = useRef(new Map<string, (ok: boolean) => void>());
  const current = useRef(conversation);
  current.current = conversation;
  const running = conversation.items.some((i) => i.type === "agent" && i.status === "running");

  // Host connection (once) and its capabilities (refreshed when they change).
  useEffect(() => {
    let connection: Host | null = null;
    const describe = async () => {
      if (!connection) return;
      const [tools, resources, prompts] = await Promise.all([
        connection.tools(),
        connection.resources(),
        connection.prompts(),
      ]);
      setHost({
        status: "connected",
        host: connection,
        title: connection.info.title,
        tools: tools.length,
        resources: resources.length,
        prompts,
      });
    };
    Host.connect(() => void describe()).then((connected) => {
      connection = connected;
      if (connected) void describe();
      else setHost({ status: "standalone" });
    });
    void settings.get().then((s) => s.model && AI_MODELS.includes(s.model) && setModelState(s.model));
  }, []);

  const refreshList = useCallback(async () => setList(await conversations.list(scope)), [scope]);

  // Once the scope is known: reopen the latest conversation of this app.
  useEffect(() => {
    if (host.status === "connecting") return;
    void conversations.list(scope).then(async (all) => {
      setList(all);
      const latest = all[0] && (await conversations.get(all[0].id));
      setConversation(latest ?? blank(scope));
    });
  }, [host.status, scope]);

  /** Applies a change to the open conversation (and persists it, throttled while running). */
  const saveTimer = useRef<ReturnType<typeof setTimeout>>(undefined);
  const update = useCallback(
    (change: (c: Conversation) => Conversation, persist: "now" | "later" = "now") => {
      const next = change(current.current);
      current.current = next;
      setConversation(next);
      clearTimeout(saveTimer.current);
      const save = () => conversations.put(current.current).then(refreshList);
      if (persist === "now") void save();
      else saveTimer.current = setTimeout(save, 800);
    },
    [refreshList],
  );

  const send = useCallback(
    async (text: string) => {
      text = text.trim();
      if (!text || running) return;
      const { approval } = await settings.get();
      const history = historyText(current.current.items);
      const trace = newStep("agent", "agent", { input: text });
      const item: AgentItem = { id: crypto.randomUUID(), type: "agent", at: Date.now(), status: "running", model, trace };
      update((c) => ({
        ...c,
        title: c.items.length === 0 ? text.slice(0, 60) : c.title,
        updatedAt: Date.now(),
        items: [...c.items, { id: crypto.randomUUID(), type: "user", text, at: Date.now() }, item],
      }));

      // The run mutates its trace in place; each change re-renders a copy of the item.
      const patch = (p: Partial<AgentItem> = {}, persist: "now" | "later" = "later") => {
        Object.assign(item, p);
        update((c) => ({ ...c, items: c.items.map((i) => (i.id === item.id ? { ...item, trace: { ...item.trace } } : i)) }), persist);
      };
      const controller = new AbortController();
      abort.current = controller;
      try {
        const outcome = await runAgent(
          {
            host: host.status === "connected" ? host.host : null,
            model,
            signal: controller.signal,
            approve:
              approval === "ask"
                ? (step: Step) => new Promise<boolean>((resolve) => approvals.current.set(step.id, resolve))
                : undefined,
            changed: () => patch(),
            setPlan: (plan) => patch({ plan }),
            addUndo: (undo) => patch({ undo: [...(item.undo ?? []), undo] }),
            stored: new Map(),
          },
          trace,
          { kind: "user", text, history },
        );
        Object.assign(trace, { status: "done", endedAt: Date.now() });
        patch({ status: "done", text: outcome.answer, question: outcome.question }, "now");
      } catch (error) {
        const stopped = controller.signal.aborted;
        closeRunning(trace, stopped ? "stopped" : message(error));
        patch({ status: stopped ? "stopped" : "error", error: stopped ? undefined : message(error) }, "now");
      } finally {
        abort.current = null;
        approvals.current.clear();
      }
    },
    [host, model, running, update],
  );

  const stop = useCallback(() => {
    abort.current?.abort();
    for (const resolve of approvals.current.values()) resolve(false);
  }, []);

  /** Answer to a pending approval. */
  const decide = useCallback((stepId: string, ok: boolean) => {
    approvals.current.get(stepId)?.(ok);
    approvals.current.delete(stepId);
  }, []);

  /** Reverts the changes of a run (host undo calls, most recent first). */
  const undo = useCallback(
    async (item: AgentItem) => {
      if (host.status !== "connected" || !item.undo?.length) return;
      for (const action of [...item.undo].reverse()) await host.host.call(action.name, action.arguments);
      update((c) => ({ ...c, items: c.items.map((i) => (i.id === item.id ? { ...i, undone: true } : i)) }));
    },
    [host, update],
  );

  return {
    host,
    conversation,
    list,
    running,
    model,
    setModel: (value: string) => {
      setModelState(value);
      void settings.set({ model: value });
    },
    send,
    stop,
    decide,
    undo,
    newConversation: () => !running && setConversation(blank(scope)),
    open: async (id: string) => {
      const found = !running && (await conversations.get(id));
      if (found) setConversation(found);
    },
    remove: async (id: string) => {
      await conversations.remove(id);
      await refreshList();
      if (id === current.current.id) setConversation(blank(scope));
    },
    rename: (title: string) => update((c) => ({ ...c, title: title.trim() || c.title })),
  };
}

/** Marks the steps left running as stopped/failed. */
function closeRunning(step: Step, reason: string) {
  if (step.status === "running" || step.status === "waiting")
    Object.assign(step, { status: "error", error: reason, endedAt: Date.now() });
  step.children.forEach((c) => closeRunning(c, reason));
}
