/**
 * The layout of every screen: the AI assistant on the left (foldable, resizable), the work area on the
 * right with one header (screen / chosen item, actions) and the screen's tabs under it. Every tab stays
 * mounted (embedded reports and the chat keep their state); only the active one is shown.
 */
import { useState, type PointerEvent, type ReactNode } from "react";
import { CheckIcon, ChevronsUpDownIcon, PanelLeftCloseIcon, PanelLeftOpenIcon, PlusIcon, SparklesIcon, type LucideIcon } from "lucide-react";
import { Chat } from "./chat";
import { Button } from "./ui/button";
import { Command, CommandEmpty, CommandGroup, CommandInput, CommandItem, CommandList, CommandSeparator } from "./ui/command";
import { Popover, PopoverContent, PopoverTrigger } from "./ui/popover";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "./ui/tabs";
import { Tooltip, TooltipContent, TooltipTrigger } from "./ui/tooltip";
import { AI_PANEL } from "@/config";
import type { Agent } from "@/agent/loop";
import { last } from "@/lib/last";
import { cn } from "@/lib/utils";

/* ---------------------------------- AI panel --------------------------------- */

const MIN = 320;
const MAX = 760;
const clamp = (w: number) => Math.round(Math.min(Math.max(w, MIN), Math.min(MAX, window.innerWidth - 480)));

export type AiSlot = {
  /** Conversations are kept per scope (e.g. "editor:<report id>"). */
  scope: string;
  agent?: Agent;
  placeholder: string;
  /** Shown instead of the chat when there is no agent (e.g. nothing chosen yet). */
  empty?: string;
  onConversation?: (id: string) => void;
};

/** The assistant: folded to a thin bar, or open and resizable (drag its right edge; width remembered). */
function AiPanel({ ai }: { ai: AiSlot }) {
  const [open, setOpen] = useState(() => last.get("panel") !== "closed");
  const [width, setWidth] = useState(() => Number(last.get("panel-width")) || 400);
  const [dragging, setDragging] = useState(false);
  const toggle = (next: boolean) => {
    last.set("panel", next ? "open" : "closed");
    setOpen(next);
  };
  const resize = (e: PointerEvent<HTMLDivElement>) => {
    e.preventDefault();
    const start = e.clientX - width;
    let next = width;
    const move = (m: globalThis.PointerEvent) => setWidth((next = clamp(m.clientX - start)));
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setDragging(false);
      last.set("panel-width", String(next));
    };
    setDragging(true);
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  const chat = ai.agent ? <Chat scope={ai.scope} agent={ai.agent} placeholder={ai.placeholder} onConversation={ai.onConversation} /> : <Empty>{ai.empty ?? "Choose an item to talk about it."}</Empty>;
  return (
    <aside style={open ? { width } : undefined} className={cn("relative flex min-h-0 shrink-0 flex-col border-r bg-sidebar", open ? "" : "w-12 items-center")}>
      {open && (
        <>
          <div
            role="separator"
            aria-orientation="vertical"
            aria-label="Resize the assistant"
            onPointerDown={resize}
            onDoubleClick={() => (setWidth(400), last.set("panel-width", "400"))}
            className={cn("absolute inset-y-0 -right-1 z-20 w-2 cursor-col-resize transition-colors hover:bg-primary/20", dragging && "bg-primary/30")}
          />
          {/* While dragging, embedded reports (iframes) must not swallow the pointer. */}
          {dragging && <div className="fixed inset-0 z-50 cursor-col-resize" />}
        </>
      )}
      <div className={cn("flex h-14 w-full shrink-0 items-center border-b", open ? "gap-2 px-4" : "justify-center")}>
        {open && (
          <span className="flex items-center gap-2 text-[14.5px] font-semibold">
            <span className="flex size-7 items-center justify-center rounded-md bg-primary/10 text-primary">
              <SparklesIcon className="size-4" />
            </span>
            Assistant
          </span>
        )}
        <Tooltip>
          <TooltipTrigger asChild>
            <Button variant="ghost" size="icon" className={cn("size-8 text-muted-foreground", open && "ml-auto")} aria-label={open ? "Fold the assistant" : "Open the assistant"} onClick={() => toggle(!open)}>
              {open ? <PanelLeftCloseIcon /> : <PanelLeftOpenIcon />}
            </Button>
          </TooltipTrigger>
          <TooltipContent side="right">{open ? "Fold the assistant" : "Open the assistant"}</TooltipContent>
        </Tooltip>
      </div>
      {!open && (
        <Button variant="ghost" size="icon" className="mt-2 size-9 text-primary" aria-label="Open the assistant" onClick={() => toggle(true)}>
          <SparklesIcon />
        </Button>
      )}
      {/* Kept mounted while folded: the conversation goes on. */}
      <div className={open ? "flex min-h-0 w-full flex-1 flex-col" : "hidden"}>{chat}</div>
    </aside>
  );
}

/* ---------------------------------- Workspace -------------------------------- */

export type WorkspaceTab = { id: string; label: string; icon?: LucideIcon; content: ReactNode; count?: number };

export function Workspace({
  title,
  icon: Icon,
  picker,
  actions,
  ai,
  tabs,
  tab,
  onTab,
  toolbar,
}: {
  title: string;
  icon: LucideIcon;
  /** The chosen item (report, dataset, file…), next to the screen's name. */
  picker?: ReactNode;
  actions?: ReactNode;
  ai?: AiSlot;
  tabs: WorkspaceTab[];
  tab?: string;
  onTab?: (id: string) => void;
  /** Right of the tabs (e.g. the tab's own buttons). */
  toolbar?: ReactNode;
}) {
  const [own, setOwn] = useState(tabs[0]?.id);
  const active = tab && tabs.some((t) => t.id === tab) ? tab : tabs.some((t) => t.id === own) ? own : tabs[0]?.id;
  const select = (id: string) => (onTab ? onTab(id) : setOwn(id));
  return (
    <div className="flex h-full min-w-0">
      {ai && AI_PANEL && <AiPanel ai={ai} />}
      <Tabs value={active} onValueChange={select} className="flex min-w-0 flex-1 flex-col gap-0">
        <header className="flex h-14 shrink-0 items-center gap-2 border-b bg-card px-5">
          <span className="flex shrink-0 items-center gap-2 text-[14.5px] font-medium text-muted-foreground">
            <Icon className="size-4" aria-hidden />
            {title}
          </span>
          {picker && <span className="shrink-0 text-muted-foreground/50">/</span>}
          {picker}
          <span className="ml-auto" />
          <div className="flex items-center gap-2">{actions}</div>
        </header>
        {tabs.length > 1 && (
          <div className="flex h-11 shrink-0 items-end gap-3 border-b bg-card px-4">
            <TabsList variant="line" className="h-11 gap-1 p-0">
              {tabs.map((t) => (
                <TabsTrigger key={t.id} value={t.id} className="relative h-11 gap-2 rounded-none border-0 px-3 text-[13.5px] data-[state=active]:text-foreground data-[state=active]:shadow-none after:absolute after:inset-x-2 after:bottom-0 after:h-0.5 after:rounded-full after:bg-primary after:opacity-0 data-[state=active]:after:opacity-100 [&_svg]:size-4">
                  {t.icon && <t.icon aria-hidden />}
                  {t.label}
                  {t.count !== undefined && <span className="rounded-full bg-muted px-1.5 text-[11px] text-muted-foreground tabular-nums">{t.count}</span>}
                </TabsTrigger>
              ))}
            </TabsList>
            {toolbar && <div className="ml-auto flex h-11 items-center gap-2">{toolbar}</div>}
          </div>
        )}
        {tabs.map((t) => (
          <TabsContent key={t.id} value={t.id} forceMount className="mt-0 flex min-h-0 flex-1 flex-col data-[state=inactive]:hidden">
            {t.content}
          </TabsContent>
        ))}
      </Tabs>
    </div>
  );
}

/** Tabs inside a tab (e.g. Schema › Tables / Columns / Measures): pills, the content scrolls under them. */
export function SubTabs({ tabs, value, onChange, toolbar }: { tabs: { id: string; label: string; count?: number; content: ReactNode }[]; value?: string; onChange?: (id: string) => void; toolbar?: ReactNode }) {
  const [own, setOwn] = useState(tabs[0]?.id);
  const active = value ?? own;
  return (
    <Tabs value={active} onValueChange={(id) => (onChange ? onChange(id) : setOwn(id))} className="flex min-h-0 flex-1 flex-col gap-0">
      <div className="flex shrink-0 flex-wrap items-center gap-3 px-6 pt-5 pb-3">
        <TabsList>
          {tabs.map((t) => (
            <TabsTrigger key={t.id} value={t.id} className="gap-1.5 px-3">
              {t.label}
              {t.count !== undefined && <span className="text-[11px] text-muted-foreground tabular-nums">{t.count}</span>}
            </TabsTrigger>
          ))}
        </TabsList>
        {toolbar && <div className="ml-auto flex items-center gap-2">{toolbar}</div>}
      </div>
      {tabs.map((t) => (
        <TabsContent key={t.id} value={t.id} className="mt-0 min-h-0 flex-1 overflow-auto px-6 pb-6">
          {t.content}
        </TabsContent>
      ))}
    </Tabs>
  );
}

/* --------------------------------- Picker ----------------------------------- */

export type PickerItem = { id: string; label: string; hint?: string };

/** The screen's item (report, dataset, file…): a searchable list, with an "Add" entry at the bottom. */
export function EntityPicker({ items, value, onChange, placeholder, onAdd, addLabel = "Add…" }: { items: PickerItem[]; value?: string; onChange: (id: string) => void; placeholder: string; onAdd?: () => void; addLabel?: string }) {
  const [open, setOpen] = useState(false);
  const current = items.find((i) => i.id === value);
  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button variant="ghost" className="h-9 max-w-xs min-w-0 gap-2 px-2 text-[14.5px] font-semibold" aria-label={placeholder}>
          <span className={cn("truncate", !current && "font-medium text-muted-foreground")}>{current?.label ?? placeholder}</span>
          {current?.hint && <span className="shrink-0 rounded bg-muted px-1.5 py-0.5 text-[11px] font-medium text-muted-foreground">{current.hint}</span>}
          <ChevronsUpDownIcon className="size-3.5 shrink-0 text-muted-foreground" />
        </Button>
      </PopoverTrigger>
      <PopoverContent align="start" className="w-[360px] p-0">
        <Command>
          <CommandInput placeholder="Search…" />
          <CommandList>
            <CommandEmpty>Nothing found.</CommandEmpty>
            {items.length > 0 && (
              <CommandGroup>
                {items.map((i) => (
                  <CommandItem key={i.id} value={`${i.label} ${i.id}`} onSelect={() => (onChange(i.id), setOpen(false))} className="gap-2">
                    <CheckIcon className={cn("size-4 text-primary", i.id === value ? "opacity-100" : "opacity-0")} />
                    <span className="min-w-0 flex-1 truncate">{i.label}</span>
                    {i.hint && <span className="text-[11.5px] text-muted-foreground">{i.hint}</span>}
                  </CommandItem>
                ))}
              </CommandGroup>
            )}
            {onAdd && (
              <>
                <CommandSeparator />
                <CommandGroup>
                  <CommandItem onSelect={() => (setOpen(false), onAdd())} className="gap-2 text-primary">
                    <PlusIcon className="size-4" /> {addLabel}
                  </CommandItem>
                </CommandGroup>
              </>
            )}
          </CommandList>
        </Command>
      </PopoverContent>
    </Popover>
  );
}

/* ---------------------------------- Helpers --------------------------------- */

export function Empty({ children, icon: Icon, action }: { children: ReactNode; icon?: LucideIcon; action?: ReactNode }) {
  return (
    <div className="m-auto flex max-w-sm flex-col items-center gap-3 p-8 text-center">
      {Icon && (
        <span className="flex size-11 items-center justify-center rounded-xl bg-muted text-muted-foreground">
          <Icon className="size-5" aria-hidden />
        </span>
      )}
      <p className="text-[13.5px] leading-relaxed text-muted-foreground">{children}</p>
      {action}
    </div>
  );
}

/** A titled block of a page (settings, schema…). */
export function Section({ title, description, actions, children, className }: { title: ReactNode; description?: ReactNode; actions?: ReactNode; children: ReactNode; className?: string }) {
  return (
    <section className={cn("rounded-xl border bg-card shadow-xs", className)}>
      <div className="flex items-start gap-3 border-b px-5 py-3.5">
        <div className="min-w-0 flex-1">
          <h2 className="text-[14.5px] font-semibold">{title}</h2>
          {description && <p className="mt-0.5 text-[12.5px] leading-relaxed text-muted-foreground">{description}</p>}
        </div>
        {actions && <div className="flex shrink-0 items-center gap-2">{actions}</div>}
      </div>
      <div className="p-5">{children}</div>
    </section>
  );
}

/** Inline status of an action: busy, done or error (next to its buttons). */
export function StatusText({ state }: { state: { busy?: string; done?: string; error?: string } }) {
  if (state.busy) return <span className="text-[12.5px] text-muted-foreground">{state.busy}</span>;
  if (state.done) return <span className="text-[12.5px] text-success">{state.done}</span>;
  if (state.error)
    return (
      <span className="max-w-md truncate text-[12.5px] text-destructive" title={state.error}>
        {state.error}
      </span>
    );
  return null;
}
