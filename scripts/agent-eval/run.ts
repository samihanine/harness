/**
 * Agent eval: runs real requests through the app's agent loop (src/agent/loop.ts + llm.ts, the real
 * model) on in-memory data, answers the assistant's questions automatically, then checks the result.
 *
 *   bun run eval:agent                  list the scenarios
 *   bun run eval:agent slides-new       run one (or several) scenarios
 *   bun run eval:agent --all            run them all
 *   options: --model gpt-6-luna  --max-calls 30  --max-tokens 200000
 *   bun run eval:live <name…>     the live scenarios (real test model: see sales.ts)
 *
 * Key: AI_KEY in .env.local (never committed). Usage is capped: past --max-calls model calls or
 * --max-tokens tokens for the whole run, every further call fails and the run stops.
 * Transcripts (every prompt and reply): .local/agent-eval/<date>/<scenario>.json
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { AI_BASE_URL, AI_MODELS } from "@/agent/llm";
import { answersText, forgetSession, runTurn, type Agent, type Choices, type Update } from "@/agent/loop";
import { SCENARIOS as OFFLINE } from "./scenarios";

/** EVAL_LIVE=1 (bun run eval:live): the live scenarios on the real test model. */
const SCENARIOS = process.env.EVAL_LIVE === "1" ? (await import("./live-scenarios")).LIVE_SCENARIOS : OFFLINE;

const MOCK = process.env.EVAL_MOCK === "1";
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
  console.log(`Scenarios (bun run ${process.env.EVAL_LIVE === "1" ? "eval:live" : "eval:agent"} <name…> | --all):`);
  for (const s of SCENARIOS) console.log(`  ${s.name.padEnd(16)} ${s.description}`);
  process.exit(0);
}
if (MOCK) process.env.AI_KEY ||= "mock";
if (!process.env.AI_KEY) {
  console.error("Missing AI_KEY: put it in .env.local (AI_KEY=…).");
  process.exit(1);
}
const chosen = (runAll ? SCENARIOS : SCENARIOS.filter((s) => names.includes(s.name))).filter((s) => !MOCK || s.mock);
const unknown = names.filter((n) => !SCENARIOS.some((s) => s.name === n));
if (unknown.length) {
  console.error(`Unknown scenario(s): ${unknown.join(", ")}`);
  process.exit(1);
}

/* --------------------------------- Budget --------------------------------- */

const usage = { calls: 0, prompt: 0, completion: 0 };
const realFetch = globalThis.fetch;
/** EVAL_MOCK=1: the replies a scenario scripts stand in for the model (see Scenario.mock); every prompt sent is kept to be checked. */
let mock: { script: string[]; next: number; prompts: string[] } | undefined;
globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
  const url = String(input instanceof Request ? input.url : input);
  if (!url.startsWith(AI_BASE_URL)) return realFetch(input, init);
  if (MOCK) {
    if (!mock) throw new Error("No mock script for this scenario.");
    const body = JSON.parse(String(init?.body ?? "{}")) as { messages?: { role: string; content: unknown }[] };
    mock.prompts.push(String(body.messages?.at(-1)?.content ?? ""));
    usage.calls++;
    const reply = mock.script[mock.next++];
    if (reply === undefined) throw new Error(`The mock script ended after ${mock.script.length} replies: the loop asked for more.`);
    return new Response(JSON.stringify({ choices: [{ message: { content: reply } }], usage: { prompt_tokens: 0, completion_tokens: 0 } }), { headers: { "Content-Type": "application/json" } });
  }
  if (usage.calls >= maxCalls || usage.prompt + usage.completion >= maxTokens) throw new Error(`Eval budget reached (${usage.calls} calls, ${usage.prompt + usage.completion} tokens): stopping.`);
  usage.calls++;
  const response = await realFetch(input, init);
  const data = await response.clone().json().catch(() => undefined);
  usage.prompt += data?.usage?.prompt_tokens ?? 0;
  usage.completion += data?.usage?.completion_tokens ?? 0;
  return response;
}) as typeof fetch;

/* ------------------------------- Conversation ------------------------------ */

/** Max question rounds per scenario (each round = a few model calls). */
const MAX_ROUNDS = 3;

/** One user request, the assistant's questions answered (scenario answers, else each first option). */
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
    if (MOCK) mock = { script: scenario.mock ?? [], next: 0, prompts: [] };
    turns = await converse(agent, scenario.request, scenario.answers);
    problems = await scenario.check();
    if (mock) {
      // What the model must have been told, on every call: never search the internet (and the reply format).
      const missing = mock.prompts.filter((p) => !/NEVER search the internet or the web/.test(p) || !/ONE raw JSON object/.test(p)).length;
      if (missing) problems.push(`${missing} prompt(s) sent without the "never search the internet" instruction or the reply format`);
      if (!/^IMPORTANT — NEVER search the internet/m.test(mock.prompts[0] ?? "")) problems.push("the first prompt does not start with the web rule");
      if (mock.next < mock.script.length) problems.push(`${mock.script.length - mock.next} scripted reply(ies) never used: the loop ended early`);
    }
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
