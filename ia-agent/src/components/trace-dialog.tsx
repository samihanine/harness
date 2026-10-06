import { useState } from "react";
import { ChevronRightIcon, CopyIcon, DownloadIcon } from "lucide-react";
import { IconButton } from "@/components/icon-button";
import { Dialog, DialogContent, DialogTitle } from "@/components/ui/dialog";
import { downloadJson, duration } from "@/lib/download";
import type { AgentItem, Step } from "@/lib/types";
import { cn } from "@/lib/utils";
import { stepIcon, stepTitle } from "./steps";

/** Every step of a run: model calls (full prompts and replies), tools, sub-agents. */
export function TraceDialog({ item, onClose }: { item: AgentItem | null; onClose: () => void }) {
  return (
    <Dialog open={!!item} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="flex h-[calc(100dvh-2rem)] max-w-[calc(100%-1rem)] flex-col gap-0 p-0 sm:max-w-3xl">
        {item && (
          <>
            <div className="flex items-center gap-2 border-b py-2 pr-10 pl-4">
              <DialogTitle className="text-[13px] font-medium">Trace</DialogTitle>
              <span className="text-[12px] text-muted-foreground">
                {count(item.trace, "llm")} model calls · {count(item.trace, "tool")} tools · {count(item.trace, "agent") - 1}{" "}
                sub-agents · {duration(item.trace.startedAt, item.trace.endedAt)} · {item.model}
              </span>
              <IconButton label="Download trace" className="ml-auto" onClick={() => downloadJson(item, `trace-${item.id}`)}>
                <DownloadIcon />
              </IconButton>
            </div>
            <div className="min-h-0 flex-1 overflow-y-auto p-2">
              <Node step={item.trace} depth={0} initiallyOpen />
            </div>
          </>
        )}
      </DialogContent>
    </Dialog>
  );
}

const count = (step: Step, kind: Step["kind"]): number =>
  (step.kind === kind ? 1 : 0) + step.children.reduce((n, c) => n + count(c, kind), 0);

function Node({ step, depth, initiallyOpen = false }: { step: Step; depth: number; initiallyOpen?: boolean }) {
  const [open, setOpen] = useState(initiallyOpen);
  const Icon = stepIcon(step);
  const hasDetails = step.input !== undefined || step.output !== undefined || !!step.error;
  return (
    <div>
      <button
        type="button"
        onClick={() => setOpen(!open)}
        className="flex w-full min-w-0 items-center gap-1.5 rounded-md px-1.5 py-1 text-left text-[12px] hover:bg-muted"
        style={{ paddingLeft: depth * 14 + 6 }}
      >
        <ChevronRightIcon
          className={cn(
            "size-3 shrink-0 text-muted-foreground transition-transform",
            open && "rotate-90",
            !hasDetails && !step.children.length && "invisible",
          )}
        />
        <Icon className="size-3.5 shrink-0 text-muted-foreground" />
        <span className="shrink-0 font-medium">{stepTitle(step)}</span>
        {step.label && <span className="truncate font-mono text-[11px] text-muted-foreground">{step.label}</span>}
        <span
          className={cn(
            "ml-auto shrink-0 pl-2 text-[11px] tabular-nums text-muted-foreground",
            step.status === "error" && "text-destructive",
          )}
        >
          {step.status === "error" ? "failed · " : step.status === "running" ? "running · " : ""}
          {duration(step.startedAt, step.endedAt)}
        </span>
      </button>
      {open && (
        <div className="animate-in fade-in-0">
          {hasDetails && (
            <div className="my-1 flex flex-col gap-1.5" style={{ paddingLeft: depth * 14 + 26 }}>
              {step.input !== undefined && <Block label={step.kind === "llm" ? "Prompt" : "Input"} value={step.input} />}
              {step.output !== undefined && <Block label={step.kind === "llm" ? "Reply" : "Output"} value={step.output} />}
              {step.error && <Block label="Error" value={step.error} error />}
            </div>
          )}
          {step.children.map((child) => (
            <Node key={child.id} step={child} depth={depth + 1} />
          ))}
        </div>
      )}
    </div>
  );
}

function Block({ label, value, error }: { label: string; value: unknown; error?: boolean }) {
  const text = typeof value === "string" ? value : JSON.stringify(value, null, 2);
  return (
    <div className="group relative rounded-md border bg-muted/40">
      <div className="flex items-center justify-between px-2 pt-1 text-[11px] text-muted-foreground">
        {label} <span className="tabular-nums">{text.length.toLocaleString()} chars</span>
      </div>
      <pre
        className={cn(
          "max-h-80 overflow-auto px-2 pb-2 font-mono text-[11px] leading-normal whitespace-pre-wrap break-words",
          error && "text-destructive",
        )}
      >
        {text}
      </pre>
      <IconButton
        label="Copy"
        size="icon-xs"
        className="absolute top-5 right-1 opacity-0 group-hover:opacity-100"
        onClick={() => void navigator.clipboard.writeText(text)}
      >
        <CopyIcon />
      </IconButton>
    </div>
  );
}
