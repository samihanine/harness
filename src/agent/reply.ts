/**
 * Reading the model replies. The model has its own system prompt and is not built for JSON,
 * so parsing is tolerant (fences, text around, small syntax slips) and every failure comes
 * with a precise message that is sent back for correction.
 */

export type Call = { name: string; args: Record<string, unknown> };

export class ReplyError extends Error {}

/** Extracts the tool calls of a reply: `{"tools":[{"name","args"}]}` (and a few close variants). */
export function parseReply(text: string): Call[] {
  const value = parseJson(text);
  const list = Array.isArray(value)
    ? value
    : isObject(value) && Array.isArray(value.tools)
      ? value.tools
      : isObject(value) && Array.isArray(value.calls)
        ? value.calls
        : isObject(value) && (value.name || value.tool)
          ? [value]
          : null;
  if (!list) throw new ReplyError('The JSON object must have a "tools" array.');
  if (list.length === 0) throw new ReplyError('"tools" is empty: call at least one tool (use "answer" to reply).');
  return list.map((call, index) => {
    if (!isObject(call)) throw new ReplyError(`tools[${index}] must be an object {"name": …, "args": {…}}.`);
    const name = call.name ?? call.tool;
    if (typeof name !== "string" || !name) throw new ReplyError(`tools[${index}] has no "name".`);
    const args = call.args ?? call.arguments ?? call.input ?? call.parameters ?? {};
    if (!isObject(args)) throw new ReplyError(`tools[${index}].args must be an object.`);
    return { name, args };
  });
}

const isObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Tries the whole text, fenced blocks, then the outermost {...}; each also after light repairs. */
function parseJson(text: string): unknown {
  const candidates = [
    text.trim(),
    ...[...text.matchAll(/```(?:json)?\s*([\s\S]*?)```/g)].map((m) => m[1].trim()),
    outermost(text),
  ].filter((c): c is string => !!c);
  let lastError = "";
  for (const candidate of candidates) {
    for (const attempt of [candidate, repair(candidate)]) {
      try {
        return JSON.parse(attempt);
      } catch (error) {
        lastError = error instanceof Error ? error.message : String(error);
      }
    }
  }
  if (!text.includes("{")) throw new ReplyError("The reply contains no JSON object.");
  throw new ReplyError(`The JSON is invalid (${lastError}).`);
}

/** The first balanced {...} or [...] block, aware of strings. */
function outermost(text: string) {
  const start = text.search(/[{[]/);
  if (start < 0) return "";
  const stack: string[] = [];
  let inString = false;
  for (let i = start; i < text.length; i++) {
    const char = text[i];
    if (inString) {
      if (char === "\\") i++;
      else if (char === '"') inString = false;
    } else if (char === '"') inString = true;
    else if (char === "{" || char === "[") stack.push(char === "{" ? "}" : "]");
    else if (char === "}" || char === "]") {
      stack.pop();
      if (stack.length === 0) return text.slice(start, i + 1);
    }
  }
  // Unclosed (truncated reply): close what is open.
  return text.slice(start) + stack.reverse().join("");
}

/** Smart quotes, trailing commas, raw line breaks inside strings. */
function repair(json: string) {
  let out = "";
  let inString = false;
  for (let i = 0; i < json.length; i++) {
    let char = json[i];
    if (!inString && (char === "“" || char === "”")) char = '"';
    if (inString) {
      if (char === "\\") {
        out += char + (json[++i] ?? "");
        continue;
      }
      if (char === '"') inString = false;
      else if (char === "\n") char = "\\n";
      else if (char === "\r") char = "";
      else if (char === "\t") char = "\\t";
    } else if (char === '"') inString = true;
    out += char;
  }
  return out.replace(/,\s*([}\]])/g, "$1");
}
