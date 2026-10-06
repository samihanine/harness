/** One node of a run's trace: the agent itself, a model call, or a tool call (possibly a sub-agent). */
export type Step = {
  id: string;
  kind: "agent" | "llm" | "tool";
  name: string;
  status: "running" | "done" | "error" | "waiting";
  /** Short human summary (tool arguments, task…). */
  label?: string;
  /** Raw input: prompt for llm steps, arguments for tools. */
  input?: unknown;
  /** Raw output: reply for llm steps, result for tools. */
  output?: unknown;
  error?: string;
  startedAt: number;
  endedAt?: number;
  children: Step[];
};

export type PlanItem = { text: string; status: "todo" | "doing" | "done" };

/** Host tool call reverting a change (given by the host in the tool result `_meta.undo`). */
export type UndoAction = { name: string; arguments: Record<string, unknown> };

export type UserItem = { id: string; type: "user"; text: string; at: number };

export type AgentItem = {
  id: string;
  type: "agent";
  at: number;
  status: "running" | "done" | "error" | "stopped";
  /** Final answer (markdown). */
  text?: string;
  /** Question asked to the user, with suggested answers. */
  question?: { text: string; options: string[] };
  error?: string;
  plan?: PlanItem[];
  undo?: UndoAction[];
  undone?: boolean;
  model: string;
  trace: Step;
};

export type Item = UserItem | AgentItem;

export type Conversation = {
  id: string;
  /** App that hosts the agent (MCP server name); conversations are listed per app. */
  scope: string;
  title: string;
  createdAt: number;
  updatedAt: number;
  items: Item[];
};
