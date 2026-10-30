/**
 * What keeps the app from working, checked at start: the AI key (missing, or refused by the AI API's
 * conversation list when llm.ts has one).
 */
import * as llm from "@/agent/llm";
import { settings } from "@/lib/store";

/** Why the assistant cannot answer, or undefined when the key is there (and accepted, when it can be checked). */
export async function aiProblem(): Promise<string | undefined> {
  const { aiKey } = await settings.get();
  if (!aiKey?.trim()) return "No AI key: the assistants cannot answer.";
  // The private API's conversation list, when llm.ts gives one: the cheapest call that checks the key.
  const list = (llm as Record<string, unknown>).getConversations;
  if (typeof list !== "function") return undefined;
  try {
    await (list as () => Promise<unknown>)();
    return undefined;
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return e instanceof llm.AiAuthError ? `The AI key is not accepted (${message})` : `The AI API cannot be reached (${message})`;
  }
}
