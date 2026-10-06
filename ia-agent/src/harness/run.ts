/**
 * The agent loop.
 *
 * Each user request is a "run" in a fresh, short-lived model session: the first message carries
 * everything needed (protocol, app instructions, tools, resource list, inline resources, a
 * compact conversation history, the request); each next message only carries tool results and
 * the resources that changed. A session growing too large is restarted with a summary of the
 * progress, so the context never overflows. Sub-agents (`delegate`) are runs of their own.
 */
import { AGENT } from "@/lib/config";
import { openSession } from "@/lib/llm";
import type { Session } from "@/lib/llm";
import type { PlanItem, Step, UndoAction } from "@/lib/types";
import type { Host, HostResource } from "./host";
import { LAST_CALL, correction, firstMessage, resourceBlock, toolTurn } from "./prompts";
import type { ToolResult } from "./prompts";
import type { Call } from "./reply";
import { ReplyError, parseReply } from "./reply";
import { FINAL, builtinTools } from "./tools";
import type { ToolDef } from "./tools";

export type RunContext = {
  host: Host | null;
  model: string;
  signal: AbortSignal;
  /** Approval of a change ("ask" mode); undefined = apply directly. */
  approve?: (step: Step) => Promise<boolean>;
  /** The trace changed (UI refresh, persistence). */
  changed: () => void;
  setPlan: (plan: PlanItem[]) => void;
  addUndo: (undo: UndoAction) => void;
  /** Large tool results of the run, readable by every agent of the run. */
  stored: Map<string, string>;
};

export type Outcome = { answer?: string; question?: { text: string; options: string[] } };

type Request = { kind: "user"; text: string; history?: string } | { kind: "task"; text: string; tools?: string[] };

export const newStep = (kind: Step["kind"], name: string, extra: Partial<Step> = {}): Step => ({
  id: crypto.randomUUID(),
  kind,
  name,
  status: "running",
  startedAt: Date.now(),
  children: [],
  ...extra,
});

export async function runAgent(ctx: RunContext, trace: Step, request: Request): Promise<Outcome> {
  const subAgent = request.kind === "task";
  const child = (kind: Step["kind"], name: string, extra: Partial<Step> = {}) => {
    const step = newStep(kind, name, extra);
    trace.children.push(step);
    ctx.changed();
    return step;
  };
  const end = (step: Step, patch: Partial<Step>) => {
    Object.assign(step, { status: "done", endedAt: Date.now() }, patch);
    ctx.changed();
  };

  // What the host offers right now.
  const host = ctx.host;
  const [hostTools, resources] = host ? await Promise.all([host.tools(), host.resources()]) : [[], []];
  const allowed = request.kind === "task" && request.tools?.length ? new Set(request.tools) : null;
  const tools: ToolDef[] = [
    ...builtinTools({ subAgent }),
    ...hostTools.filter((t) => !allowed || allowed.has(t.name)),
  ];
  const byName = new Map(tools.map((t) => [t.name, t]));

  // Inline resources: sent first, then resent when they change.
  const seen = new Map<string, string>();
  const readInline = async () => {
    const changed: { resource: HostResource; text: string }[] = [];
    for (const resource of resources.filter((r) => r.inline)) {
      const text = await host!.read(resource.uri).catch((e: unknown) => `Unreadable: ${message(e)}`);
      if (text.length > AGENT.inlineChars) resource.inline = false;
      else if (seen.get(resource.uri) !== text) changed.push({ resource, text });
      if (resource.inline) seen.set(resource.uri, text);
    }
    return changed.map(({ resource, text }) => resourceBlock(resource.uri, text));
  };

  const log: string[] = [];
  const first = async (progress?: string) =>
    firstMessage({
      appTitle: host?.info.title ?? "App",
      instructions: host?.info.instructions ?? "",
      tools,
      resources,
      inline: await readInline(),
      history: request.kind === "user" ? request.history : undefined,
      request,
      progress,
    });

  let session: Session = openSession(ctx.model, subAgent ? "sub-agent" : "agent");
  let prompt = await first();
  let retries = 0;
  const maxSteps = subAgent ? AGENT.subAgentSteps : AGENT.maxSteps;

  try {
    for (let turn = 0; turn <= maxSteps; turn++) {
      ctx.signal.throwIfAborted();
      const last = turn === maxSteps;
      const llm = child("llm", "model", { input: last ? LAST_CALL : prompt });
      const reply = await session.send(last ? LAST_CALL : prompt, ctx.signal).catch((error: unknown) => {
        end(llm, { status: "error", error: message(error) });
        throw error;
      });
      end(llm, { output: reply });

      let calls: Call[];
      try {
        calls = parseReply(reply);
        const unknown = calls.filter((c) => !byName.has(c.name)).map((c) => c.name);
        if (unknown.length) throw new ReplyError(`Unknown tool(s): ${unknown.join(", ")}. Available: ${[...byName.keys()].join(", ")}.`);
        retries = 0;
      } catch (error) {
        if (!(error instanceof ReplyError)) throw error;
        llm.error = error.message;
        llm.label = "invalid reply, sent back for correction";
        // A model that keeps writing prose: show its text rather than failing.
        if (++retries > AGENT.maxRetries || last) return { answer: reply.trim() };
        prompt = correction(error.message);
        turn--;
        continue;
      }

      const final = calls.find((c) => FINAL.has(c.name));
      if (final && calls.length === 1) {
        const step = child("tool", final.name, { input: final.args });
        end(step, { output: "" });
        if (final.name === "ask_user")
          return { question: { text: String(final.args.question ?? ""), options: (final.args.options as string[]) ?? [] } };
        return { answer: String(final.args.text ?? "") };
      }
      if (last) return { answer: reply.trim() };

      const notes: string[] = [];
      if (final) notes.push(`"${final.name}" was ignored: it must be the only call of its reply. Call it alone once the other results are in.`);
      const toRun = calls.filter((c) => !FINAL.has(c.name));
      const results = await execute(toRun);
      for (const [i, r] of results.entries())
        log.push(`- ${toRun[i].name}(${summary(toRun[i].args)}) → ${r.ok ? "ok" : "error"}: ${r.text.slice(0, 160)}`);

      const updated = await readInline();
      if (session.chars > AGENT.sessionChars) {
        // Restart with a compact state instead of letting the context overflow.
        await session.close();
        session = openSession(ctx.model, subAgent ? "sub-agent" : "agent");
        seen.clear();
        prompt = await first(progress(log, results));
        child("llm", "context restarted", { status: "done", label: `${log.length} actions summarized` });
        continue;
      }
      prompt = toolTurn({ results, updated, stepsLeft: maxSteps - turn - 1, notes });
    }
    return {};
  } finally {
    await session.close();
  }

  /** Runs the calls in order; consecutive read-only calls run in parallel. */
  async function execute(calls: Call[]): Promise<ToolResult[]> {
    const results: ToolResult[] = [];
    for (let i = 0; i < calls.length; ) {
      let j = i + 1;
      if (byName.get(calls[i].name)!.readOnly) while (j < calls.length && byName.get(calls[j].name)!.readOnly) j++;
      results.push(...(await Promise.all(calls.slice(i, j).map(run))));
      i = j;
    }
    return results;
  }

  async function run(call: Call): Promise<ToolResult> {
    const tool = byName.get(call.name)!;
    const step = child("tool", call.name, { input: call.args, label: summary(call.args) });
    try {
      let text = await invoke(tool, call.args, step);
      if (text.length > AGENT.resultChars) {
        const uri = `result://${ctx.stored.size + 1}`;
        ctx.stored.set(uri, text);
        const lines = text.split("\n").length;
        text = `${text.slice(0, AGENT.resultChars / 2)}\n…\n[Truncated: ${text.length} characters, ${lines} lines. Full result stored as ${uri}: use read/search on it.]`;
      }
      end(step, { output: step.output ?? text });
      return { name: call.name, ok: true, text };
    } catch (error) {
      if (ctx.signal.aborted) throw error;
      end(step, { status: "error", error: message(error) });
      return { name: call.name, ok: false, text: message(error) };
    }
  }

  async function invoke(tool: ToolDef, args: Record<string, unknown>, step: Step): Promise<string> {
    switch (tool.builtin ? tool.name : "") {
      case "read": {
        const lines = (await text(String(args.uri))).split("\n");
        const offset = Math.max(0, Number(args.offset ?? 0));
        const limit = Math.max(1, Number(args.limit ?? AGENT.readLines));
        const slice = lines.slice(offset, offset + limit);
        const more = offset + slice.length < lines.length ? ` (more: offset ${offset + slice.length})` : "";
        return `[lines ${offset}–${offset + slice.length - 1} of ${lines.length}${more}]\n${slice.join("\n")}`;
      }
      case "search": {
        const query = String(args.query ?? "");
        const regex = /^\/(.+)\/([a-z]*)$/.exec(query);
        const test = regex
          ? ((r) => (line: string) => r.test(line))(new RegExp(regex[1], regex[2].includes("i") ? regex[2] : `${regex[2]}i`))
          : (line: string) => line.toLowerCase().includes(query.toLowerCase());
        const hits = (await text(String(args.uri)))
          .split("\n")
          .map((line, n) => [n, line] as const)
          .filter(([, line]) => test(line));
        const limit = Number(args.limit ?? 50);
        return hits.length === 0
          ? "No match."
          : `${hits.length} matching line(s)${hits.length > limit ? `, first ${limit}` : ""}:\n${hits
              .slice(0, limit)
              .map(([n, line]) => `${n}: ${line}`)
              .join("\n")}`;
      }
      case "plan": {
        const items = (args.items as PlanItem[]) ?? [];
        if (!subAgent) ctx.setPlan(items);
        return "Plan updated.";
      }
      case "delegate":
        return delegate((args.tasks as { task: string; tools?: string[] }[]) ?? [], step);
    }
    if (!host) throw new Error("No app connected.");
    if (!tool.readOnly && ctx.approve) {
      step.status = "waiting";
      ctx.changed();
      const ok = await ctx.approve(step);
      step.status = "running";
      if (!ok) throw new Error("The user declined this change.");
    }
    const outcome = await host.call(tool.name, args);
    if (outcome.isError) throw new Error(outcome.text);
    if (outcome.undo) ctx.addUndo(outcome.undo);
    return outcome.text;
  }

  async function text(uri: string) {
    const stored = ctx.stored.get(uri);
    if (stored !== undefined) return stored;
    if (!host) throw new Error(`Unknown resource ${uri}`);
    return host.read(uri);
  }

  async function delegate(tasks: { task: string; tools?: string[] }[], parent: Step) {
    if (tasks.length === 0) throw new Error("No task given.");
    const reports: string[] = new Array(tasks.length);
    let next = 0;
    const worker = async () => {
      while (next < tasks.length) {
        const index = next++;
        const sub = newStep("agent", `sub-agent ${index + 1}`, { label: tasks[index].task.slice(0, 120), input: tasks[index] });
        parent.children.push(sub);
        ctx.changed();
        try {
          const outcome = await runAgent(ctx, sub, { kind: "task", text: tasks[index].task, tools: tasks[index].tools });
          reports[index] = outcome.answer ?? "(no report)";
          Object.assign(sub, { status: "done", endedAt: Date.now(), output: reports[index] });
        } catch (error) {
          if (ctx.signal.aborted) throw error;
          reports[index] = `Failed: ${message(error)}`;
          Object.assign(sub, { status: "error", endedAt: Date.now(), error: message(error) });
        }
        ctx.changed();
      }
    };
    await Promise.all(Array.from({ length: Math.min(AGENT.maxParallel, tasks.length) }, worker));
    return reports.map((report, i) => `## Sub-agent ${i + 1}\n${report}`).join("\n\n");
  }
}

function progress(log: string[], results: ToolResult[]) {
  const recent = log.slice(-40);
  return [
    "The conversation was restarted to save space. What you already did (most recent last):",
    log.length > recent.length ? `(${log.length - recent.length} earlier actions omitted)` : "",
    ...recent,
    "",
    "Results of your last tool calls:",
    ...results.map((r) => `- ${r.name} (${r.ok ? "ok" : "error"}): ${r.text.slice(0, 1500)}`),
    "",
    "Continue the task from here.",
  ]
    .filter(Boolean)
    .join("\n");
}

/** One-line summary of tool arguments for the trace and the progress log. */
export function summary(args: Record<string, unknown> | undefined) {
  if (!args) return "";
  return Object.entries(args)
    .map(([key, value]) => `${key}=${typeof value === "string" ? value : JSON.stringify(value)}`)
    .join(", ")
    .slice(0, 140);
}

export const message = (error: unknown) => (error instanceof Error ? error.message : String(error));
