# ia-agent

AI agent panel, 100% front-end, embedded by other apps in an `<iframe>`.
Apps plug in with the **Model Context Protocol**: the host page runs an MCP server
(official SDK) and connects it to the iframe with the official `PostMessageTransport`.

```bash
bun install
bun dev            # http://localhost:3200 — playground: http://localhost:3200/#/playground
bun run export-html   # exports/ia-agent.html (single file)
bun run export-src    # exports/ia-agent-src.zip
```

## Plugging an app in

```ts
import { McpServer } from "@modelcontextprotocol/server";
import { PostMessageTransport } from "@modelcontextprotocol/ext-apps";
import { z } from "zod";

const server = new McpServer(
  { name: "my-app", title: "My app", version: "1.0.0" },  // name = conversation scope
  { instructions: "What the agent must know about this app." },
);
server.registerResource("rows", "table://rows", { mimeType: "application/jsonl", annotations: { priority: 1 } },
  async (uri) => ({ contents: [{ uri: uri.href, text: rowsAsJsonLines() }] }));
server.registerTool("update_row", { description: "…", inputSchema: z.object({ id: z.string(), … }) },
  async (args) => ({ content: [{ type: "text", text: "ok" }], _meta: { undo: { name: "update_row", arguments: before } } }));
server.registerPrompt("summary", { title: "Summarize" }, () => ({ messages: [{ role: "user", content: { type: "text", text: "Summarize…" } }] }));

const target = iframe.contentWindow!;
await server.connect(new PostMessageTransport(target, target));
```

See `src/playground/demo-server.ts` for a complete example.

### Conventions on top of MCP

| What | How |
| --- | --- |
| Context always sent | resource with `annotations.priority ≥ 0.8` — sent with each request, resent when it changes |
| Large data | other resources / templates — listed only, the agent reads them by lines (`read`, `search`) |
| Undo | tool result `_meta.undo = { name, arguments }` (host call that reverts the change) |
| No approval needed | tool `annotations.readOnlyHint: true` (also run in parallel) |
| Suggestions | MCP prompts (shown in the empty conversation) |
| App instructions | server `instructions` |

Prefer one JSON object per line (JSONL) for tabular resources: the agent pages and searches by lines.

## Harness

- `src/lib/llm.ts`: plain-text LLM access; provider constants at the top (stateless or stateful API, SSE).
- `src/harness/run.ts`: the loop. One short-lived model session per request; first message = protocol,
  app instructions, tools, resource list, inline resources, compact history, request; then only tool
  results and changed resources. Large results are stored (`result://n`) and read on demand. A session
  growing too large is restarted with a summary of the progress. `delegate` runs sub-agents in parallel.
- `src/harness/reply.ts`: tolerant JSON parsing; invalid replies are sent back with the precise error.
- `src/lib/config.ts`: limits (steps, sizes, sub-agents…).
- Every run keeps a full trace (prompts, replies, tools, sub-agents): "View trace" on each answer;
  conversations (with traces) export as JSON.

## Evaluation

`evals/` runs the real harness on realistic host apps (MCP servers in memory, same tools as excel-agent /
pbi-agent) with a real model, and checks the outcome (data changed correctly, question asked when it should…):

```bash
bun test evals --timeout 900000                  # AI_KEY from .env, model gpt-6-luna
EVAL_MODEL=gpt-5 EVAL_ONLY=pbi bun test evals --timeout 900000
```

Full traces are written to `evals/out/` (gitignored).
