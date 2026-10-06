/** Exact model names sent with each request (the first one is the default). */
export const MODELS = ["gpt-6-luna", "gpt-5.5", "gpt-5", "gpt-5-mini"] as const;

/** Harness tuning. Sizes are in characters (≈ 4 characters per token). */
export const AGENT = {
  /** Tool turns allowed for one user request before the agent must answer. */
  maxSteps: 20,
  /** Invalid replies sent back for correction before the raw text is shown as the answer. */
  maxRetries: 2,
  /** Tool results longer than this are stored; the model gets a preview and reads the rest. */
  resultChars: 4_000,
  /** Inline resources (priority ≥ 0.8) larger than this are only listed, to be read on demand. */
  inlineChars: 16_000,
  /** Previous exchanges replayed at the start of a request. */
  historyChars: 12_000,
  /** A session holding more than this is restarted with a compact summary of the progress. */
  sessionChars: 160_000,
  /** Lines returned by `read` when no limit is given. */
  readLines: 200,
  /** Sub-agents (`delegate` tool). */
  subAgents: true,
  /** Sub-agents running at the same time. */
  maxParallel: 3,
  /** Tool turns allowed for a sub-agent. */
  subAgentSteps: 12,
};
