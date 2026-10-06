import { AlertCircleIcon, CheckIcon, CircleIcon, CopyIcon, ListTreeIcon, LoaderIcon, Undo2Icon } from "lucide-react";
import { IconButton } from "@/components/icon-button";
import { Button } from "@/components/ui/button";
import type { AgentItem, Item, PlanItem } from "@/lib/types";
import { cn } from "@/lib/utils";
import { Markdown } from "./markdown";
import { RunSteps } from "./steps";

type Actions = {
  onDecide: (stepId: string, ok: boolean) => void;
  onTrace: (item: AgentItem) => void;
  onUndo: (item: AgentItem) => void;
  onReply: (text: string) => void;
  canReply: boolean;
};

export function Message({ item, ...actions }: { item: Item } & Actions) {
  if (item.type === "user")
    return (
      <div className="max-w-[88%] self-end rounded-xl rounded-br-sm bg-muted px-3 py-2 whitespace-pre-wrap animate-in fade-in-0 slide-in-from-bottom-1">
        {item.text}
      </div>
    );
  return <AgentMessage item={item} {...actions} />;
}

function AgentMessage({ item, onDecide, onTrace, onUndo, onReply, canReply }: { item: AgentItem } & Actions) {
  const copy = item.text ?? item.question?.text ?? "";
  return (
    <div className="group/message flex flex-col gap-2 animate-in fade-in-0">
      <RunSteps item={item} onDecide={onDecide} />
      {item.plan && item.plan.length > 0 && <Plan items={item.plan} />}
      {item.text && <Markdown>{item.text}</Markdown>}
      {item.question && (
        <div className="flex flex-col gap-2">
          <Markdown>{item.question.text}</Markdown>
          {item.question.options.length > 0 && (
            <div className="flex flex-wrap gap-1.5">
              {item.question.options.map((option) => (
                <Button key={option} variant="outline" size="sm" disabled={!canReply} onClick={() => onReply(option)}>
                  {option}
                </Button>
              ))}
            </div>
          )}
        </div>
      )}
      {item.status === "error" && (
        <p className="flex items-start gap-1.5 rounded-md border border-destructive/20 bg-destructive/5 px-2.5 py-1.5 text-[12px] text-destructive">
          <AlertCircleIcon className="mt-0.5 size-3.5 shrink-0" /> {item.error}
        </p>
      )}
      {item.status === "stopped" && <p className="text-[12px] text-muted-foreground">Stopped.</p>}
      {item.status !== "running" && (
        <div className="-ml-1.5 flex items-center gap-0.5 opacity-0 transition-opacity group-hover/message:opacity-100 has-[[data-pinned]]:opacity-100">
          {copy && (
            <IconButton label="Copy" size="icon-xs" onClick={() => void navigator.clipboard.writeText(copy)}>
              <CopyIcon />
            </IconButton>
          )}
          <IconButton label="View trace" size="icon-xs" onClick={() => onTrace(item)}>
            <ListTreeIcon />
          </IconButton>
          {!!item.undo?.length && !item.undone && (
            <IconButton label={`Undo ${item.undo.length} change(s)`} size="icon-xs" data-pinned onClick={() => onUndo(item)}>
              <Undo2Icon />
            </IconButton>
          )}
          {item.undone && <span className="px-1 text-[11px] text-muted-foreground">Changes undone</span>}
          <span className="ml-1 text-[11px] text-muted-foreground/70">{item.model}</span>
        </div>
      )}
    </div>
  );
}

function Plan({ items }: { items: PlanItem[] }) {
  return (
    <ul className="flex flex-col gap-1 rounded-lg border bg-card px-3 py-2 shadow-soft">
      {items.map((step, i) => (
        <li key={i} className={cn("flex items-start gap-2 text-[12px]", step.status === "done" && "text-muted-foreground line-through")}>
          {step.status === "done" ? (
            <CheckIcon className="mt-0.5 size-3.5 shrink-0 text-success" />
          ) : step.status === "doing" ? (
            <LoaderIcon className="mt-0.5 size-3.5 shrink-0 animate-spin" />
          ) : (
            <CircleIcon className="mt-0.5 size-3.5 shrink-0 text-muted-foreground/60" />
          )}
          {step.text}
        </li>
      ))}
    </ul>
  );
}
