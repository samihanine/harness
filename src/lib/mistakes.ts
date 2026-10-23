/**
 * The AI's past mistakes, kept in this browser: each failed tool call (or unreadable reply) with the
 * call that fixed it in the same turn, if any. The latest ones of a screen's tools are shown to the AI
 * as examples not to follow; Settings downloads them all (to improve prompts and tools).
 */
import { get, set } from "idb-keyval";

export type Mistake = { at: string; scope: string; tool: string; args: string; error: string; fix?: string };

const KEY = "mistakes";
const MAX = 500;

export const allMistakes = async (): Promise<Mistake[]> => (await get(KEY)) ?? [];
export const clearMistakes = () => set(KEY, []);

let queue: Promise<unknown> = Promise.resolve();
export const addMistakes = (items: Mistake[]) => {
  if (!items.length) return queue;
  return (queue = queue.then(async () => set(KEY, [...(await allMistakes()), ...items].slice(-MAX))).catch(() => undefined));
};

/**
 * The latest mistakes of these tools (those fixed first), one per tool and error, as short lines; empty
 * when there are none.
 */
export async function mistakesText(tools: string[], max = 5) {
  const names = new Set(tools);
  const seen = new Set<string>();
  const picked = (await allMistakes().catch(() => []))
    .filter((m) => names.has(m.tool))
    .reverse()
    .filter((m) => {
      const key = `${m.tool}:${m.error.slice(0, 60)}`;
      return !seen.has(key) && seen.add(key);
    })
    .sort((a, b) => Number(!!b.fix) - Number(!!a.fix))
    .slice(0, max);
  return picked.map((m) => `- ${m.tool}(${m.args.slice(0, 220)}) → ERROR: ${m.error.split("\n")[0].slice(0, 220)}${m.fix ? `\n  worked instead: ${m.tool}(${m.fix.slice(0, 220)})` : ""}`).join("\n");
}
