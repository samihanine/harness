/**
 * Prompts of the harness. Every message to the model is written by the application, not by a
 * person: it is made of tagged sections and always ends with the expected reply format.
 */
import type { HostResource } from "./host";
import type { ToolDef } from "./tools";
import { signature } from "./tools";

export const tag = (name: string, body: string, attrs = "") => `<${name}${attrs ? ` ${attrs}` : ""}>\n${body.trim()}\n</${name}>`;

const PROTOCOL = `You are the engine of an AI agent embedded in an application. This message is written by the APPLICATION, not by a person: it gives you the context and the user's request, runs the tools you call, and sends you their results. The user only sees what you pass to the "answer" tool.

HOW YOU WORK
1. Every reply you write is ONE raw JSON object, and nothing else: no prose before or after, no markdown fences, no comments.
2. The JSON object lists the tools to run now:
   {"tools": [{"name": "<tool>", "args": {<arguments>}}]}
3. You may call several independent tools in one reply. They run in order; you get all results in the next message.
4. When you have everything you need, reply with the "answer" tool ALONE. It ends your turn.
5. Never invent data: read it with tools. Never claim a change you did not make through a tool.
6. Keep going until the request is fully done (or truly blocked), then answer. Do not stop to describe what you are about to do.

GOOD PRACTICES
- Large content is not pasted here: it is listed in <resources>; read what you need with "read" (by lines) or "search".
- Tool results longer than a few thousand characters are stored as result://… and shown as a preview; read them if needed.
- For multi-step work, keep a short checklist with "plan".
- When a tool fails, read the error, fix the arguments and retry, or explain the problem in your answer.
- Write the answer in the user's language, concise, in markdown.

EXAMPLES OF VALID REPLIES
{"tools": [{"name": "read", "args": {"uri": "table://orders", "offset": 0, "limit": 50}}]}
{"tools": [{"name": "plan", "args": {"items": [{"text": "Read the data", "status": "doing"}, {"text": "Update rows", "status": "todo"}]}}, {"name": "read", "args": {"uri": "table://orders"}}]}
{"tools": [{"name": "answer", "args": {"text": "Done: **3 rows** updated."}}]}`;

export const REPLY_FORMAT = tag(
  "reply_format",
  `Reply now with ONE raw JSON object and nothing else: {"tools": [{"name": …, "args": {…}}]}.
When you are done, your final reply is still JSON: {"tools": [{"name": "answer", "args": {"text": "…"}}]}`,
);

export function toolsSection(tools: ToolDef[]) {
  return tag("tools", tools.map(signature).join("\n"));
}

export function resourcesSection(resources: HostResource[]) {
  if (resources.length === 0) return "";
  const lines = resources.map(
    (r) =>
      `- ${r.uri}${r.template ? " (template)" : ""}: ${r.name}${r.description ? `. ${r.description}` : ""}${r.inline ? " [included below]" : r.lines ? ` [${r.lines} lines, too large to include: use search, or read a range]` : ""}`,
  );
  return tag("resources", lines.join("\n"));
}

export const resourceBlock = (uri: string, content: string) => tag("resource", content, `uri="${uri}"`);

export type FirstMessage = {
  appTitle: string;
  instructions: string;
  tools: ToolDef[];
  resources: HostResource[];
  inline: string[];
  history?: string;
  request: { kind: "user"; text: string } | { kind: "task"; text: string };
  /** Progress carried over when a long session is restarted. */
  progress?: string;
};

export function firstMessage(m: FirstMessage) {
  return [
    tag("protocol", PROTOCOL),
    m.request.kind === "task" &&
      tag(
        "role",
        "You are a SUB-AGENT working for an orchestrator agent. Do the task below with your tools, then call answer with a complete report (the orchestrator only sees this report).",
      ),
    m.instructions && tag("app_instructions", m.instructions, `name="${m.appTitle}"`),
    toolsSection(m.tools),
    resourcesSection(m.resources),
    m.inline.length > 0 && tag("inline_resources", m.inline.join("\n")),
    m.history && tag("conversation_so_far", m.history),
    m.progress && tag("progress_so_far", m.progress),
    m.request.kind === "user" ? tag("user_message", m.request.text) : tag("task", m.request.text),
    REPLY_FORMAT,
  ]
    .filter(Boolean)
    .join("\n\n");
}

export type ToolResult = { name: string; ok: boolean; text: string };

export function toolTurn({
  results,
  updated,
  stepsLeft,
  notes,
}: {
  results: ToolResult[];
  updated: string[];
  stepsLeft: number;
  notes: string[];
}) {
  const body = results
    .map((r, i) => tag("result", r.text || "(empty)", `n="${i + 1}" tool="${r.name}" status="${r.ok ? "ok" : "error"}"`))
    .join("\n");
  return [
    tag("tool_results", body),
    updated.length > 0 && tag("updated_resources", `These resources changed since you last saw them:\n${updated.join("\n")}`),
    ...notes.map((note) => tag("note", note)),
    stepsLeft <= 3 && tag("note", `Only ${stepsLeft} tool turn(s) left: finish the essential, then answer.`),
    REPLY_FORMAT,
  ]
    .filter(Boolean)
    .join("\n\n");
}

export const correction = (error: string) =>
  [
    tag(
      "invalid_reply",
      `Your previous reply could not be used: ${error}\nThe application only understands one raw JSON object. Do not apologize, just send the corrected JSON.`,
    ),
    REPLY_FORMAT,
  ].join("\n\n");

export const LAST_CALL = [
  tag("note", "No tool turns left. Answer now with what you have, and say what remains to be done."),
  REPLY_FORMAT,
].join("\n\n");
