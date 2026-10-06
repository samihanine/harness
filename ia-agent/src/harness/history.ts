/** Compact replay of the previous exchanges, given at the start of each request. */
import { AGENT } from "@/lib/config";
import type { Item, Step } from "@/lib/types";

const clip = (text: string, max: number) => (text.length > max ? `${text.slice(0, max)}…` : text);

/** Tools used by a run (main agent level), e.g. "read ×3, update_rows". */
function actions(trace: Step) {
  const counts = new Map<string, number>();
  for (const step of trace.children)
    if (step.kind === "tool" && !["answer", "ask_user", "plan"].includes(step.name))
      counts.set(step.name, (counts.get(step.name) ?? 0) + 1);
  return [...counts].map(([name, n]) => (n > 1 ? `${name} ×${n}` : name)).join(", ");
}

export function historyText(items: Item[]) {
  const lines: string[] = [];
  let size = 0;
  let index = items.length - 1;
  for (; index >= 0; index--) {
    const item = items[index];
    let line: string;
    if (item.type === "user") line = `USER: ${clip(item.text, 2000)}`;
    else {
      const done = actions(item.trace);
      const body = item.question
        ? `asked: ${item.question.text}${item.question.options.length ? ` (options: ${item.question.options.join(" / ")})` : ""}`
        : item.text
          ? clip(item.text, 1500)
          : `(${item.status}${item.error ? `: ${item.error}` : ""})`;
      line = `ASSISTANT${done ? ` [used: ${done}]` : ""}${item.undone ? " [its changes were undone]" : ""}: ${body}`;
    }
    if (size + line.length > AGENT.historyChars) break;
    lines.unshift(line);
    size += line.length;
  }
  if (lines.length === 0) return "";
  return [index >= 0 && `(${index + 1} earlier messages omitted)`, ...lines].filter(Boolean).join("\n\n");
}
