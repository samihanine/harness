/** The provider chosen in Settings › AI, behind one interface (server/llm/types.ts). */
import type { Settings } from "../settings";
import { anthropicProvider } from "./anthropic";
import { maiaChat, MAIA_MODELS } from "./maia";
import { AuthError, type ChatProvider } from "./types";

export function providerFor(s: Settings, conversationId: () => string): ChatProvider {
  if (s.ai.provider === "maia") return maiaChat(s.ai.maia, conversationId);
  if (!s.ai.anthropicKey) throw new AuthError("No Anthropic API key: add one in Settings › AI (or set CLAUDE_API_KEY).");
  return anthropicProvider(s.ai.anthropicKey);
}

export const ANTHROPIC_MODELS = ["claude-opus-5-5", "claude-sonnet-5-5", "claude-haiku-5-5", "claude-fable-5-2"];
/** Claude Code takes the models its proxy offers; these are the usual ones. */
export const CLAUDE_CODE_MODELS = ["claude-opus-5-5", "claude-sonnet-5-5", "claude-haiku-5-5"];

export const modelsOf = (provider: Settings["ai"]["provider"]) => (provider === "maia" ? MAIA_MODELS : provider === "claude-code" ? CLAUDE_CODE_MODELS : ANTHROPIC_MODELS);
