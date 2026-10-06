/**
 * Example host: what an app does to plug the agent in. It is a plain MCP server
 * (official SDK) connected to the agent iframe with the official PostMessageTransport.
 */
import { McpServer, ResourceTemplate } from "@modelcontextprotocol/server";
import { z } from "zod";

export type Task = { id: string; title: string; status: "todo" | "doing" | "done" };

export function createDemoServer(store: { get: () => Task[]; set: (tasks: Task[]) => void }) {
  const server = new McpServer(
    { name: "playground", title: "Playground", version: "1.0.0" },
    { instructions: "The user manages a small task list. Keep titles short. Statuses: todo, doing, done." },
  );
  const text = (value: unknown) => ({ content: [{ type: "text" as const, text: JSON.stringify(value) }] });

  // Inline resource (priority 1): sent with every request, resent when it changes. One task per line.
  server.registerResource(
    "tasks",
    "tasks://all",
    { title: "Tasks", description: "One JSON task per line", mimeType: "application/jsonl", annotations: { priority: 1 } },
    async (uri) => ({ contents: [{ uri: uri.href, text: store.get().map((t) => JSON.stringify(t)).join("\n") }] }),
  );
  server.registerResource(
    "task",
    new ResourceTemplate("tasks://{id}", { list: undefined }),
    { title: "One task", mimeType: "application/json" },
    async (uri, { id }) => ({ contents: [{ uri: uri.href, text: JSON.stringify(store.get().find((t) => t.id === id) ?? null) }] }),
  );

  server.registerTool(
    "add_task",
    {
      description: "Adds a task.",
      inputSchema: z.object({ title: z.string(), status: z.enum(["todo", "doing", "done"]).optional() }),
    },
    async ({ title, status }) => {
      const task: Task = { id: crypto.randomUUID().slice(0, 8), title, status: status ?? "todo" };
      store.set([...store.get(), task]);
      return { ...text(task), _meta: { undo: { name: "delete_task", arguments: { id: task.id } } } };
    },
  );
  server.registerTool(
    "update_task",
    {
      description: "Changes the title and/or status of a task.",
      inputSchema: z.object({
        id: z.string(),
        title: z.string().optional(),
        status: z.enum(["todo", "doing", "done"]).optional(),
      }),
    },
    async ({ id, ...patch }) => {
      const before = store.get().find((t) => t.id === id);
      if (!before) return { ...text(`No task ${id}`), isError: true };
      store.set(store.get().map((t) => (t.id === id ? { ...t, ...patch } : t)));
      return { ...text("ok"), _meta: { undo: { name: "update_task", arguments: { id, title: before.title, status: before.status } } } };
    },
  );
  server.registerTool(
    "delete_task",
    { description: "Deletes a task.", inputSchema: z.object({ id: z.string() }) },
    async ({ id }) => {
      const before = store.get().find((t) => t.id === id);
      if (!before) return { ...text(`No task ${id}`), isError: true };
      store.set(store.get().filter((t) => t.id !== id));
      return { ...text("deleted"), _meta: { undo: { name: "add_task", arguments: { title: before.title, status: before.status } } } };
    },
  );
  server.registerPrompt("plan-week", { title: "Organize my tasks by priority" }, () => ({
    messages: [{ role: "user", content: { type: "text", text: "Organize my tasks by priority and mark the first one as doing." } }],
  }));
  return server;
}
