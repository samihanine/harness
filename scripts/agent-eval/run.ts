/**
 * Agent eval: runs real requests through the app's agent loop (src/agent/loop.ts + llm.ts, the real
 * model) on in-memory data, answers the planner's questions automatically, then checks the result.
 *
 *   bun run eval:agent                  list the scenarios
 *   bun run eval:agent slides-new       run one (or several) scenarios
 *   bun run eval:agent --all            run them all
 *   options: --model gpt-6-luna  --max-calls 30  --max-tokens 200000
 *
 * Key: AI_KEY in .env.local (never committed). Usage is capped: past --max-calls model calls or
 * --max-tokens tokens for the whole run, every further call fails and the run stops.
 * Transcripts (every prompt and reply): .local/agent-eval/<date>/<scenario>.json
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { AI_BASE_URL, AI_MODELS } from "@/agent/llm";
import { answersText, forgetSession, runTurn, type Agent, type Choices, type Update } from "@/agent/loop";
import { SCENARIOS } from "./scenarios";

const args = process.argv.slice(2);
const option = (name: string, fallback: string) => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args.splice(i, 2)[1] : fallback;
};
const model = option("model", process.env.AI_MODEL || AI_MODELS[0]);
const maxCalls = Number(option("max-calls", "30"));
const maxTokens = Number(option("max-tokens", "200000"));
const runAll = args.includes("--all");
const names = args.filter((a) => !a.startsWith("--"));

if (!runAll && !names.length) {
  console.log("Scenarios (bun run eval:agent <name…> | --all):");
  for (const s of SCENARIOS) console.log(`  ${s.name.padEnd(16)} ${s.description}`);
  process.exit(0);
}
if (!process.env.AI_KEY) {
  console.error("Missing AI_KEY: put it in .env.local (AI_KEY=…).");
  process.exit(1);
}
const chosen = runAll ? SCENARIOS : SCENARIOS.filter((s) => names.includes(s.name));
const unknown = names.filter((n) => !SCENARIOS.some((s) => s.name === n));
if (unknown.length) {
  console.error(`Unknown scenario(s): ${unknown.join(", ")}`);
  process.exit(1);
}

/* --------------------------------- Budget --------------------------------- */

const usage = { calls: 0, prompt: 0, completion: 0 };
const realFetch = globalThis.fetch;
globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
  const url = String(input instanceof Request ? input.url : input);
  if (!url.startsWith(AI_BASE_URL)) return realFetch(input, init);
  if (usage.calls >= maxCalls || usage.prompt + usage.completion >= maxTokens) throw new Error(`Eval budget reached (${usage.calls} calls, ${usage.prompt + usage.completion} tokens): stopping.`);
  usage.calls++;
  const response = await realFetch(input, init);
  const data = await response.clone().json().catch(() => undefined);
  usage.prompt += data?.usage?.prompt_tokens ?? 0;
  usage.completion += data?.usage?.completion_tokens ?? 0;
  return response;
}) as typeof fetch;

/* ------------------------------- Conversation ------------------------------ */

/** Max planner question rounds per scenario (each round = a few model calls). */
const MAX_ROUNDS = 3;

/** One user request, the planner's questions answered (scenario answers, else each first option). */
async function converse(agent: Agent, request: string, answers: string[] = []) {
  const conversationId = crypto.randomUUID();
  const turns: { request: string; update: Update }[] = [];
  let message = request;
  let answering = false;
  let history = "";
  try {
    for (let round = 0; round <= MAX_ROUNDS; round++) {
      let update: Update = { steps: [], llm: [] };
      for await (const u of runTurn({ agent, model, conversationId, history, request: message, answering, signal: AbortSignal.timeout(180_000) })) update = u;
      turns.push({ request: message, update });
      const asked = update.steps.find((s) => s.name === "ask_user")?.args as Choices | undefined;
      if (!asked) break;
      if (round === MAX_ROUNDS) throw new Error(`Still asking after ${MAX_ROUNDS} rounds of answers.`);
      history += `user: ${message}\nassistant asked: ${asked.questions.map((q) => `${q.question} [${q.options.join(" | ")}]`).join(" / ")}\n`;
      message = answers[round] ?? answersText(asked, asked.questions.map((q) => [q.options[0]]));
      answering = true;
    }
  } finally {
    forgetSession(conversationId);
  }
  return turns;
}

/* ----------------------------------- Run ----------------------------------- */

const dir = `.local/agent-eval/${new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19)}`;
mkdirSync(dir, { recursive: true });
console.log(`Model ${model} · budget ${maxCalls} calls / ${maxTokens} tokens · transcripts in ${dir}\n`);
let failed = 0;

for (const scenario of chosen) {
  const before = { ...usage };
  const t0 = performance.now();
  let turns: Awaited<ReturnType<typeof converse>> = [];
  let problems: string[] = [];
  let error: string | undefined;
  try {
    const agent = await scenario.setup();
    turns = await converse(agent, scenario.request, scenario.answers);
    problems = await scenario.check();
  } catch (e) {
    error = e instanceof Error ? e.message : String(e);
  }
  const steps = turns.flatMap((t) => t.update.steps);
  const toolErrors = steps.filter((s) => s.error);
  const ok = !error && !problems.length;
  if (!ok) failed++;
  const spent = { calls: usage.calls - before.calls, tokens: usage.prompt + usage.completion - before.prompt - before.completion };
  console.log(`${ok ? "✓" : "✗"} ${scenario.name} — ${turns.length} turn(s), ${spent.calls} model calls, ${spent.tokens} tokens, ${Math.round((performance.now() - t0) / 1000)} s`);
  console.log(`  tools: ${steps.map((s) => `${s.name}${s.error ? "✗" : ""}`).join(", ") || "none"}`);
  if (toolErrors.length) console.log(`  tool errors:\n${toolErrors.map((s) => `    - ${s.name}: ${String(s.result).slice(0, 200)}`).join("\n")}`);
  if (error) console.log(`  ERROR: ${error}`);
  for (const p of problems) console.log(`  ✗ ${p}`);
  const answer = turns.at(-1)?.update.answer;
  if (answer) console.log(`  answer: ${answer.replace(/\s+/g, " ").slice(0, 240)}`);
  console.log();
  writeFileSync(`${dir}/${scenario.name}.json`, JSON.stringify({ scenario: scenario.name, model, ok, error, problems, spent, turns }, null, 2));
  if (error?.startsWith("Eval budget reached")) break;
}

console.log(`${chosen.length - failed}/${chosen.length} passed · total ${usage.calls} model calls, ${usage.prompt} prompt + ${usage.completion} completion tokens`);
process.exit(failed ? 1 : 0);
