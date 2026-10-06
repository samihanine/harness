/** Built-in tools of the harness, always available next to the host tools. */
import { AGENT } from "@/lib/config";
import type { JsonSchema } from "./host";

export type ToolDef = {
  name: string;
  description: string;
  schema: JsonSchema;
  readOnly: boolean;
  builtin?: boolean;
};

const str = (description: string): JsonSchema => ({ type: "string", description });
const int = (description: string): JsonSchema => ({ type: "integer", description });
const object = (properties: Record<string, JsonSchema>, required: string[]): JsonSchema => ({
  type: "object",
  properties,
  required,
});

/** Ends the turn. */
export const FINAL = new Set(["answer", "ask_user"]);

export function builtinTools({ subAgent }: { subAgent: boolean }): ToolDef[] {
  const tools: ToolDef[] = [
    {
      name: "answer",
      description: subAgent
        ? "Final report of your task for the orchestrator: findings, what you changed, open problems. Ends your work. Must be the only call of its turn."
        : "Final reply to the user, in markdown, in the user's language. Ends the turn. Must be the only call of its turn.",
      schema: object({ text: str("The reply") }, ["text"]),
      readOnly: true,
    },
    {
      name: "read",
      description:
        "Reads a resource (or a stored tool result `result://…`) by lines. Use offset/limit to page through large ones.",
      schema: object(
        {
          uri: str("Resource URI; fill the {variables} of templates"),
          offset: int("First line, 0-based (default 0)"),
          limit: int(`Number of lines (default ${AGENT.readLines})`),
        },
        ["uri"],
      ),
      readOnly: true,
    },
    {
      name: "search",
      description: "Lines of a resource (or stored result) matching a query: case-insensitive text, or /regex/.",
      schema: object({ uri: str("Resource URI"), query: str("Text or /regex/"), limit: int("Max lines (default 50)") }, [
        "uri",
        "query",
      ]),
      readOnly: true,
    },
    {
      name: "plan",
      description:
        "Creates or replaces your visible checklist for multi-step work. Update it as you progress (one item `doing` at a time).",
      schema: object(
        {
          items: {
            type: "array",
            items: object({ text: str("Step"), status: { type: "string", enum: ["todo", "doing", "done"] } }, [
              "text",
              "status",
            ]),
          },
        },
        ["items"],
      ),
      readOnly: true,
    },
  ];
  if (subAgent) return tools;
  tools.splice(1, 0, {
    name: "ask_user",
    description:
      "Asks the user a question when a decision is really theirs (ambiguous request, destructive choice). Ends the turn. Must be the only call of its turn.",
    schema: object(
      { question: str("The question"), options: { type: "array", items: str("Suggested answer"), description: "2–4 short suggested answers" } },
      ["question"],
    ),
    readOnly: true,
  });
  if (AGENT.subAgents)
    tools.push({
      name: "delegate",
      description: `Runs sub-agents in parallel (max ${AGENT.maxParallel} at once), each on a self-contained task, and returns their reports. Use for independent heavy work (exploring several sources, building several parts). Sub-agents do not see this conversation: put every needed detail in the task.`,
      schema: object(
        {
          tasks: {
            type: "array",
            items: object(
              {
                task: str("Complete, self-contained instructions and the expected report"),
                tools: { type: "array", items: str("Tool name"), description: "Host tools allowed (default: all)" },
              },
              ["task"],
            ),
          },
        },
        ["tasks"],
      ),
      readOnly: false,
    });
  return tools.map((tool) => ({ ...tool, builtin: true }));
}

/** Compact TypeScript-like signature of a tool, e.g. `read(uri: string, offset?: integer)`. */
export function signature(tool: ToolDef) {
  const props = tool.schema.properties ?? {};
  const required = new Set(tool.schema.required ?? []);
  const args = Object.entries(props).map(([name, schema]) => `${name}${required.has(name) ? "" : "?"}: ${typeOf(schema)}`);
  const notes = Object.entries(props)
    .filter(([, schema]) => schema.description)
    .map(([name, schema]) => `    ${name}: ${schema.description}`);
  return [`- ${tool.name}(${args.join(", ")})${tool.readOnly ? "" : "  [changes data]"}`, `    ${tool.description}`, ...notes].join(
    "\n",
  );
}

function typeOf(schema: JsonSchema, depth = 0): string {
  if (!schema || typeof schema !== "object") return "any";
  if (schema.enum) return schema.enum.map((v) => JSON.stringify(v)).join(" | ");
  if (schema.anyOf) return schema.anyOf.map((s) => typeOf(s, depth)).join(" | ");
  const type = Array.isArray(schema.type) ? schema.type.filter((t) => t !== "null").join(" | ") : schema.type;
  if (type === "array") return `${wrap(typeOf(schema.items ?? {}, depth + 1))}[]`;
  if (type === "object" || schema.properties) {
    if (!schema.properties || depth > 2) return "object";
    const required = new Set(schema.required ?? []);
    const fields = Object.entries(schema.properties).map(
      ([name, s]) => `${name}${required.has(name) ? "" : "?"}: ${typeOf(s, depth + 1)}`,
    );
    return `{ ${fields.join("; ")} }`;
  }
  return type ?? "any";
}

const wrap = (type: string) => (type.includes("|") ? `(${type})` : type);
