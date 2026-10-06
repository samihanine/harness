import { useState } from "react";
import {
  BotIcon,
  CheckIcon,
  ChevronRightIcon,
  FileTextIcon,
  HandIcon,
  ListChecksIcon,
  SearchIcon,
  SparklesIcon,
  WrenchIcon,
  XIcon,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { duration } from "@/lib/download";
import type { AgentItem, Step } from "@/lib/types";
import { cn } from "@/lib/utils";

const ICONS: Record<string, typeof FileTextIcon> = {
  read: FileTextIcon,
  search: SearchIcon,
  plan: ListChecksIcon,
  delegate: BotIcon,
  model: SparklesIcon,
};

export const stepIcon = (step: Step) => (step.kind === "agent" ? BotIcon : (ICONS[step.name] ?? WrenchIcon));

export const stepTitle = (step: Step) =>
  step.kind === "agent"
    ? step.name
    : step.kind === "llm"
      ? step.name === "model"
        ? "Model call"
        : step.name
      : ({ read: "Read", search: "Search", plan: "Plan", delegate: "Delegate", answer: "Answer", ask_user: "Question" }[
          step.name
        ] ?? step.name.replace(/[_-]+/g, " "));

/** Steps shown in the chat: tool calls of the main agent (not the final answer). */
const visible = (trace: Step) =>
  trace.children.filter((s) => s.kind === "tool" && !["answer", "ask_user"].includes(s.name));

export function RunSteps({ item, onDecide }: { item: AgentItem; onDecide: (id: string, ok: boolean) => void }) {
  const [open, setOpen] = useState(false);
  const steps = visible(item.trace);
  const running = item.status === "running";
  const waiting = steps.some((s) => s.status === "waiting");
  const thinking = running && item.trace.children.at(-1)?.kind === "llm";

  if (!running && steps.length === 0) return null;
  const expanded = running || open;
  return (
    <div className="flex flex-col gap-0.5 text-muted-foreground">
      {!running && (
        <button
          type="button"
          onClick={() => setOpen(!open)}
          className="-ml-1 flex w-fit items-center gap-1 rounded px-1 py-0.5 text-[12px] transition-colors hover:text-foreground"
        >
          <ChevronRightIcon className={cn("size-3 transition-transform", open && "rotate-90")} />
          {steps.length} step{steps.length > 1 ? "s" : ""} · {duration(item.trace.startedAt, item.trace.endedAt)}
        </button>
      )}
      {expanded && (
        <div className={cn("flex flex-col gap-0.5", !running && "ml-1 border-l pl-2.5 animate-in fade-in-0 slide-in-from-top-1")}>
          {steps.map((step) => (
            <StepLine key={step.id} step={step} onDecide={onDecide} />
          ))}
        </div>
      )}
      {thinking && !waiting && <span className="shimmer w-fit text-[12px]">Thinking…</span>}
    </div>
  );
}

function StepLine({ step, onDecide }: { step: Step; onDecide: (id: string, ok: boolean) => void }) {
  const Icon = stepIcon(step);
  const subAgents = step.children.filter((c) => c.kind === "agent");
  return (
    <div className="animate-in fade-in-0 slide-in-from-bottom-1 duration-200">
      <div className="flex min-w-0 items-center gap-1.5 py-0.5 text-[12px]">
        <Icon className="size-3.5 shrink-0 opacity-70" />
        <span className={cn("shrink-0 font-medium text-foreground/80", step.status === "running" && "shimmer")}>
          {stepTitle(step)}
        </span>
        {step.label && <span className="truncate font-mono text-[11px] opacity-70">{step.label}</span>}
        <span className="ml-auto flex shrink-0 items-center gap-1 pl-2">
          {step.status === "error" && <XIcon className="size-3 text-destructive" />}
          {step.status === "done" && <CheckIcon className="size-3 opacity-50" />}
        </span>
      </div>
      {step.status === "error" && step.error && (
        <p className="ml-5 line-clamp-2 text-[11px] text-destructive/80">{step.error}</p>
      )}
      {step.status === "waiting" && (
        <div className="my-1 ml-5 flex items-center gap-1.5 rounded-md border bg-card p-1.5 pl-2.5 shadow-soft animate-in fade-in-0 zoom-in-95">
          <HandIcon className="size-3.5" />
          <span className="text-[12px] text-foreground">Allow this change?</span>
          <Button size="xs" variant="ghost" className="ml-auto" onClick={() => onDecide(step.id, false)}>
            Deny
          </Button>
          <Button size="xs" onClick={() => onDecide(step.id, true)}>
            Allow
          </Button>
        </div>
      )}
      {subAgents.length > 0 && (
        <div className="ml-1.5 flex flex-col border-l pl-2.5">
          {subAgents.map((agent) => (
            <div key={agent.id} className="flex min-w-0 items-center gap-1.5 py-0.5 text-[12px]">
              <BotIcon className="size-3 shrink-0 opacity-60" />
              <span className={cn("truncate", agent.status === "running" && "shimmer")}>{agent.label ?? agent.name}</span>
              <span className="ml-auto shrink-0 pl-2 text-[11px] opacity-60">
                {agent.children.filter((c) => c.kind === "tool").length} tools
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
