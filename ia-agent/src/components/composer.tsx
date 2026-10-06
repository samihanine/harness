import { useRef, useState } from "react";
import { ArrowUpIcon, ChevronDownIcon, SquareIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { AI_MODELS } from "@/lib/llm";

export function Composer({
  running,
  model,
  onModel,
  onSend,
  onStop,
  draft,
  onDraft,
}: {
  running: boolean;
  model: string;
  onModel: (model: string) => void;
  onSend: (text: string) => void;
  onStop: () => void;
  draft: string;
  onDraft: (text: string) => void;
}) {
  const ref = useRef<HTMLTextAreaElement>(null);
  const [focused, setFocused] = useState(false);
  const submit = () => {
    if (running || !draft.trim()) return;
    onSend(draft);
    onDraft("");
  };
  return (
    <div
      data-focused={focused || undefined}
      className="rounded-xl border bg-card shadow-soft transition-shadow data-focused:border-ring/60 data-focused:shadow-float"
      onClick={() => ref.current?.focus()}
    >
      <textarea
        ref={ref}
        rows={1}
        value={draft}
        placeholder="Ask anything…"
        onFocus={() => setFocused(true)}
        onBlur={() => setFocused(false)}
        onChange={(e) => onDraft(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
            e.preventDefault();
            submit();
          }
        }}
        className="field-sizing-content max-h-60 min-h-11 w-full resize-none bg-transparent px-3 pt-2.5 outline-none placeholder:text-muted-foreground/70"
      />
      <div className="flex items-center gap-1 px-1.5 pb-1.5">
        <DropdownMenu>
          <DropdownMenuTrigger
            render={<Button variant="ghost" size="xs" className="gap-0.5 text-muted-foreground hover:text-foreground" />}
          >
            {model} <ChevronDownIcon className="size-3" />
          </DropdownMenuTrigger>
          <DropdownMenuContent className="w-40">
            {AI_MODELS.map((m) => (
              <DropdownMenuItem key={m} onClick={() => onModel(m)}>
                {m}
              </DropdownMenuItem>
            ))}
          </DropdownMenuContent>
        </DropdownMenu>
        <span className="ml-auto" />
        {running ? (
          <Button size="icon-sm" variant="secondary" className="rounded-full" aria-label="Stop" onClick={onStop}>
            <SquareIcon className="size-3 fill-current" />
          </Button>
        ) : (
          <Button size="icon-sm" className="rounded-full transition-transform active:scale-95" aria-label="Send" disabled={!draft.trim()} onClick={submit}>
            <ArrowUpIcon />
          </Button>
        )}
      </div>
    </div>
  );
}
