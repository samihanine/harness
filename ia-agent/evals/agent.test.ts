/**
 * Live evaluation of the harness on realistic host apps (real model, real MCP in memory).
 *   AI_KEY=… bun test evals --timeout 600000      (AI_KEY is read from .env)
 *   EVAL_MODEL=gpt-5 EVAL_ONLY=dax bun test evals …
 */
import { expect, mock, test } from "bun:test";
import { mkdirSync, writeFileSync } from "node:fs";

mock.module("@/lib/store", () => ({
  settings: { get: async () => ({ aiKey: process.env.AI_KEY ?? "", model: "", approval: "auto" }), set: async () => {} },
}));

const { Client, InMemoryTransport } = await import("@modelcontextprotocol/client");
const { Host } = await import("@/harness/host");
const { newStep, runAgent } = await import("@/harness/run");
const { SCENARIOS } = await import("./scenarios");
import type { Step } from "@/lib/types";

const MODEL = process.env.EVAL_MODEL ?? "gpt-6-luna";
const count = (s: Step, kind: Step["kind"]): number => (s.kind === kind ? 1 : 0) + s.children.reduce((n, c) => n + count(c, kind), 0);
const invalid = (s: Step): number => (s.kind === "llm" && s.error ? 1 : 0) + s.children.reduce((n, c) => n + invalid(c), 0);
const errors = (s: Step): string[] => [...(s.kind === "tool" && s.status === "error" ? [`${s.name}: ${s.error}`] : []), ...s.children.flatMap(errors)];
const chars = (s: Step): number => (s.kind === "llm" ? String(s.input ?? "").length + String(s.output ?? "").length : 0) + s.children.reduce((n, c) => n + chars(c), 0);

test("scenarios", async () => {
  expect(process.env.AI_KEY).toBeTruthy();
  const only = process.env.EVAL_ONLY;
  const results = await Promise.all(
    SCENARIOS.filter((s) => !only || s.name.includes(only)).map(async (scenario) => {
      const { server, check } = scenario.server();
      const [a, b] = InMemoryTransport.createLinkedPair();
      await server.connect(a);
      const client = new Client({ name: "agent", version: "1" });
      await client.connect(b);
      const trace = newStep("agent", "agent", { input: scenario.prompt });
      const start = Date.now();
      let outcome = {};
      let crash = "";
      try {
        outcome = await runAgent(
          { host: Host.fromClient(client), model: MODEL, signal: new AbortController().signal, changed: () => {}, setPlan: () => {}, addUndo: () => {}, stored: new Map() },
          trace,
          { kind: "user", text: scenario.prompt, history: scenario.history },
        );
      } catch (e) {
        crash = e instanceof Error ? e.message : String(e);
      }
      trace.endedAt = Date.now();
      const problems = crash ? [`crash: ${crash}`] : check(outcome);
      return { name: scenario.name, ok: problems.length === 0, problems, seconds: (Date.now() - start) / 1000, llm: count(trace, "llm"), tools: count(trace, "tool"), invalid: invalid(trace), toolErrors: errors(trace), chars: chars(trace), outcome, trace };
    }),
  );
  mkdirSync("evals/out", { recursive: true });
  writeFileSync(`evals/out/${MODEL}-${Date.now()}.json`, JSON.stringify(results, null, 1));
  for (const r of results)
    console.log(
      `${r.ok ? "PASS" : "FAIL"} ${r.name} | ${r.seconds.toFixed(0)}s, ${r.llm} model calls (${r.invalid} invalid), ${r.tools} tools, ${Math.round(r.chars / 1000)}k chars${r.problems.length ? `\n   - ${r.problems.join("\n   - ")}` : ""}${r.toolErrors.length ? `\n   tool errors: ${r.toolErrors.join(" | ").slice(0, 400)}` : ""}`,
    );
});
