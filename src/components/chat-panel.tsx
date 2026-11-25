/**
 * The assistant of a screen: its conversation (answers in markdown, the steps it took, questions with
 * buttons), the mascot dressed for the screen, and the composer (text, pictures and files).
 */
import { useEffect, useRef, useState, type ReactNode } from "react";
import { ArrowUpIcon, CheckIcon, ChevronRightIcon, CircleAlertIcon, DownloadIcon, EllipsisIcon, ExternalLinkIcon, FolderOpenIcon, HistoryIcon, Loader2Icon, MessageSquareTextIcon, PaperclipIcon, PanelLeftCloseIcon, PlusIcon, RefreshCwIcon, SquareIcon, ThumbsDownIcon, ThumbsUpIcon, Trash2Icon, XIcon } from "lucide-react";
import { toast } from "sonner";
import type { ChatItem, FileRef, ScreenId } from "@shared/types";
import { Markdown } from "./markdown";
import { Button } from "./ui/button";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuSeparator, DropdownMenuTrigger } from "./ui/dropdown-menu";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "./ui/dialog";
import { refreshProject, useProjectsStatus } from "@/lib/projects";
import { Mascot, outfitOf } from "@/features/mascot";
import type { useChat } from "@/lib/chat";
import { errorText, fileUrl } from "@/lib/api";
import { keepFile, useProjects } from "@/lib/store";
import { useSettings } from "@/lib/settings";
import { cn } from "@/lib/utils";

type Chat = ReturnType<typeof useChat>;
type Assistant = Extract<ChatItem, { kind: "assistant" }>;

const LABELS: Record<ScreenId, string> = { trainer: "Trainer", builder: "Builder", scientist: "Data Scientist", presenter: "Presenter" };

/** The projects whose Confluence pages are the context (global choice, several allowed). */
function ProjectsMenu() {
  const settings = useSettings();
  const [chosen, setChosen] = useProjects();
  const [open, setOpen] = useState(false);
  const status = useProjectsStatus(open);
  const [busy, setBusy] = useState(false);
  const all = settings?.confluence.projects ?? [];
  if (!all.length) return null;
  const ids = chosen ?? all.slice(0, 1).map((p) => p.id);
  const names = all.filter((p) => ids.includes(p.id)).map((p) => p.name);
  return (
    <DropdownMenu open={open} onOpenChange={setOpen}>
      <DropdownMenuTrigger asChild>
        <button className="flex max-w-36 items-center gap-1 rounded-full border px-2.5 py-1 text-[12px] text-muted-foreground hover:bg-muted hover:text-foreground" title="Projects: their Confluence pages are the assistant's context">
          <FolderOpenIcon className="size-3.5 shrink-0" />
          <span className="truncate">{names.length ? names.join(", ") : "No project"}</span>
        </button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end" className="w-64">
        <div className="px-2 py-1.5 text-[12px] text-muted-foreground">Projects: their Confluence pages (read once and kept) are given to the assistant at the start of a conversation.</div>
        {all.map((p) => {
          const on = ids.includes(p.id);
          const st = status?.find((x) => x.id === p.id);
          return (
            <DropdownMenuItem key={p.id} onSelect={(e) => (e.preventDefault(), setChosen(on ? ids.filter((x) => x !== p.id) : [...ids, p.id]))}>
              <span className={cn("flex size-4 items-center justify-center rounded border", on && "border-foreground bg-foreground text-background")}>{on && <CheckIcon className="size-3" />}</span>
              <span className="flex-1 truncate">{p.name}</span>
              {st && <span className="text-[11px] text-muted-foreground">{st.error ? "!" : `${st.pages} p.`}</span>}
              {st?.url && (
                <a href={st.url} target="_blank" rel="noreferrer" onClick={(e) => e.stopPropagation()} title="Open the folder in Confluence" className="text-muted-foreground hover:text-foreground">
                  <ExternalLinkIcon className="size-3.5" />
                </a>
              )}
            </DropdownMenuItem>
          );
        })}
        <DropdownMenuSeparator />
        <DropdownMenuItem disabled={busy || !ids.length} onSelect={(e) => (e.preventDefault(), setBusy(true), Promise.all(ids.map(refreshProject)).then(() => toast.success("Pages read again: the next conversations get them."), (x) => toast.error(errorText(x))).finally(() => setBusy(false)))}>
          {busy ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh the chosen projects' pages
        </DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

/** The open conversation's feedback: thumbs, a comment, the JSON file. */
function Feedback({ chat }: { chat: Chat }) {
  const c = chat.conversation;
  const [commenting, setCommenting] = useState(false);
  const [text, setText] = useState("");
  if (!c) return null;
  return (
    <>
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <Button variant="ghost" size="icon-sm" aria-label="This conversation">
            {c.rating === "up" ? <ThumbsUpIcon className="text-success" /> : c.rating === "down" ? <ThumbsDownIcon className="text-destructive" /> : <EllipsisIcon />}
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="end" className="w-60">
          <div className="px-2 py-1.5 text-[12px] text-muted-foreground">How was this conversation?</div>
          <div className="flex gap-1 px-2 pb-1.5">
            <Button variant={c.rating === "up" ? "dark" : "outline"} size="sm" className="flex-1" onClick={() => void chat.rate({ rating: c.rating === "up" ? undefined : "up" })}><ThumbsUpIcon /> Good</Button>
            <Button variant={c.rating === "down" ? "dark" : "outline"} size="sm" className="flex-1" onClick={() => void chat.rate({ rating: c.rating === "down" ? undefined : "down" })}><ThumbsDownIcon /> Bad</Button>
          </div>
          <DropdownMenuItem onSelect={() => (setText(c.comment ?? ""), setCommenting(true))}><MessageSquareTextIcon /> {c.comment ? "Edit the comment" : "Add a comment"}</DropdownMenuItem>
          <DropdownMenuItem onSelect={chat.exportJson}><DownloadIcon /> Download (JSON)</DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
      <Dialog open={commenting} onOpenChange={setCommenting}>
        <DialogContent className="max-w-md rounded-3xl">
          <DialogHeader>
            <DialogTitle>Comment on this conversation</DialogTitle>
          </DialogHeader>
          <textarea autoFocus value={text} onChange={(e) => setText(e.target.value)} rows={5} placeholder="What went well, what went wrong…" className="input h-auto py-2" />
          <Button variant="dark" onClick={() => void chat.rate({ comment: text.trim() || undefined }).then(() => setCommenting(false))}>Save</Button>
        </DialogContent>
      </Dialog>
    </>
  );
}

export function ChatPanel({ chat, screen, placeholder, disabled, accept = "image/*", suggestions, onFold, itemLabel }: { chat: Chat; screen: ScreenId; placeholder?: string; disabled?: string; accept?: string; suggestions?: string[]; onFold?: () => void; /** The name of a conversation's item (history). */ itemLabel?: (item: string) => string | undefined }) {
  const settings = useSettings();
  const items = chat.conversation?.items ?? [];
  const scroller = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = scroller.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [items.length, items.at(-1)?.kind === "assistant" ? (items.at(-1) as Assistant).text.length : 0, chat.status]);
  const mascot = settings?.general.mascot !== false;
  const lastAssistant = items.findLastIndex((i) => i.kind === "assistant");
  const state = chat.running ? (chat.status?.includes("Saving") || chat.status?.includes("Working") ? "working" : items.at(-1)?.kind === "assistant" && (items.at(-1) as Assistant).text ? "talking" : "thinking") : "idle";
  return (
    <div className="flex h-full flex-col">
      <header className="flex h-16 shrink-0 items-center gap-2 border-b px-4">
        <div className="flex size-9 items-center justify-center overflow-hidden rounded-xl bg-muted">{mascot ? <Mascot outfit={outfitOf(screen)} state={chat.running ? "thinking" : "idle"} size={34} /> : <span className="text-[13px] font-semibold">AI</span>}</div>
        <div className="min-w-0 flex-1">
          <div className="text-[14px] font-semibold">{LABELS[screen]}</div>
          <div className="truncate text-[12px] text-muted-foreground">{chat.conversation?.title ?? "New conversation"}</div>
        </div>
        <ProjectsMenu />
        <Feedback chat={chat} />
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button variant="ghost" size="icon-sm" aria-label="Conversations">
              <HistoryIcon />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="max-h-[70vh] w-80 overflow-y-auto">
            {chat.conversations.length === 0 && <div className="px-2 py-1.5 text-[13px] text-muted-foreground">No conversation yet.</div>}
            {chat.conversations.slice(0, 60).map((c) => (
              <DropdownMenuItem key={c.id} onSelect={() => chat.select(c.id, c.item)} className={cn("group", c.id === chat.conversation?.id && "bg-accent")}>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-1.5 truncate">
                    {chat.isRunning(c.id) && <Loader2Icon className="size-3 shrink-0 animate-spin" />}
                    <span className="truncate">{c.title}</span>
                  </div>
                  <div className="truncate text-[11.5px] text-muted-foreground">
                    {itemLabel?.(c.item) ?? "—"} · {new Date(c.updatedAt).toLocaleDateString(undefined, { day: "numeric", month: "short" })}
                    {c.rating === "up" ? " · 👍" : c.rating === "down" ? " · 👎" : ""}
                  </div>
                </div>
                <button
                  className="opacity-0 group-hover:opacity-100"
                  onClick={(e) => {
                    e.stopPropagation();
                    void chat.deleteConversation(c.id);
                  }}
                  aria-label="Delete"
                >
                  <Trash2Icon className="size-3.5" />
                </button>
              </DropdownMenuItem>
            ))}
            <DropdownMenuSeparator />
            <DropdownMenuItem onSelect={chat.newConversation}>
              <PlusIcon /> New conversation
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
        <Button variant="ghost" size="icon-sm" onClick={chat.newConversation} aria-label="New conversation">
          <PlusIcon />
        </Button>
        {onFold && (
          <Button variant="ghost" size="icon-sm" onClick={onFold} aria-label="Fold the assistant">
            <PanelLeftCloseIcon />
          </Button>
        )}
      </header>
      <div ref={scroller} className="min-h-0 flex-1 overflow-y-auto px-4 py-4">
        {items.length === 0 ? (
          <div className="flex h-full flex-col items-center justify-center gap-4 text-center">
            {mascot && <Mascot outfit={outfitOf(screen)} state="idle" size={110} />}
            <div className="max-w-[300px] text-[13.5px] text-muted-foreground">{disabled ?? "Ask anything: I explore, ask you to confirm, then act."}</div>
            {!disabled && suggestions?.length ? (
              <div className="flex max-w-[340px] flex-col gap-1.5">
                {suggestions.map((s) => (
                  <button key={s} onClick={() => chat.send(s)} className="rounded-2xl border bg-card px-3.5 py-2 text-left text-[13px] hover:bg-muted">
                    {s}
                  </button>
                ))}
              </div>
            ) : null}
          </div>
        ) : (
          <div className="flex flex-col gap-5">
            {items.map((it, i) =>
              it.kind === "user" ? (
                <UserBubble key={i} text={it.text} files={it.files} />
              ) : (
                <AssistantMessage key={i} item={it} last={i === lastAssistant} status={i === items.length - 1 ? chat.status : undefined} mascot={mascot && i === lastAssistant ? <Mascot outfit={outfitOf(screen)} state={state} size={30} /> : undefined} onAnswer={(t) => chat.send(t)} disabled={chat.running} />
              ),
            )}
          </div>
        )}
      </div>
      <Composer chat={chat} placeholder={placeholder} disabled={disabled} accept={accept} />
    </div>
  );
}

function UserBubble({ text, files }: { text: string; files?: FileRef[] }) {
  return (
    <div className="flex flex-col items-end gap-1.5">
      {files?.length ? (
        <div className="flex flex-wrap justify-end gap-1.5">
          {files.map((f) => (f.type.startsWith("image/") ? <img key={f.id} src={fileUrl(f.id)} alt={f.name} className="max-h-40 rounded-xl border" /> : <span key={f.id} className="rounded-lg border bg-muted px-2 py-1 text-[12px]">{f.name}</span>))}
        </div>
      ) : null}
      {text && <div className="max-w-[88%] rounded-2xl rounded-br-md bg-muted px-3.5 py-2 text-[13.5px] whitespace-pre-wrap">{text}</div>}
    </div>
  );
}

const TOOL_NAMES: Record<string, string> = {
  ask_user: "Question",
  model_structure: "Read the model",
  describe_table: "Read a table",
  search_model: "Searched the model",
  run_dax: "DAX query",
  read_view: "Read the report",
  set_slicer: "Set a slicer",
  set_filter: "Set a filter",
  go_to_page: "Changed page",
  reset_view: "Reset the report",
  visual_data: "Read a visual's data",
  look_at_report: "Looked at the report",
  report_structure: "Read a report's structure",
  list_reports: "Listed reports",
  report_overview: "Read the draft",
  add_page: "Added a page",
  add_visual: "Added a visual",
  update_visual: "Changed a visual",
  format_visual: "Formatted a visual",
  format_options: "Format options",
  layout_page: "Laid out a page",
  remove: "Removed",
  save_report: "Saved the report",
  discard_changes: "Discarded changes",
  read_definition_file: "Read a definition file",
  write_definition_file: "Wrote a definition file",
  save_measure: "Saved a measure",
  delete_measure: "Deleted a measure",
  list_sources: "Listed sources",
  load_dax: "Loaded model data",
  load_file: "Loaded a file",
  sql: "SQL query",
  create_step: "Saved a step",
  set_block: "Updated the analysis",
  remove_block: "Removed a block",
  read_analysis: "Read the analysis",
  read_deck: "Read the deck",
  write_deck: "Wrote the deck",
  edit_slides: "Edited slides",
  look_at_deck: "Looked at the slides",
  capture_powerbi: "Captured Power BI",
  pptx_structure: "Read the PowerPoint",
  look_at_slides: "Looked at the slides",
  view_picture: "Looked at a picture",
  set_text: "Changed a text",
  replace_picture: "Replaced a picture",
  add_picture: "Added a picture",
  add_text: "Added a text",
  shape: "Moved a shape",
  slide: "Changed slides",
  read_page: "Read the page",
  write_page: "Wrote the page",
  read_context: "Read context pages",
};

function Steps({ tools }: { tools: Assistant["tools"] }) {
  const [open, setOpen] = useState<string>();
  const [all, setAll] = useState(false);
  const steps = tools.filter((t) => t.name !== "ask_user");
  if (!steps.length) return null;
  // Long runs: the last steps only, the others behind a line.
  const hidden = all ? 0 : Math.max(0, steps.length - 4);
  const shown = steps.slice(hidden);
  return (
    <div className="flex flex-col gap-0.5">
      {hidden > 0 && (
        <button onClick={() => setAll(true)} className="flex items-center gap-2 rounded-lg px-1.5 py-1 text-left text-[12.5px] text-muted-foreground hover:bg-muted">
          <ChevronRightIcon className="size-3.5" /> {hidden} earlier steps ({steps.filter((t) => t.status === "error").length ? `${steps.filter((t) => t.status === "error").length} failed, retried` : "all done"})
        </button>
      )}
      {shown.map((t) => (
        <div key={t.id} className="text-[12.5px]">
          <button onClick={() => setOpen(open === t.id ? undefined : t.id)} className="flex w-full items-center gap-2 rounded-lg px-1.5 py-1 text-left text-muted-foreground hover:bg-muted">
            {t.status === "running" ? <Loader2Icon className="size-3.5 animate-spin" /> : t.status === "error" ? <CircleAlertIcon className="size-3.5 text-destructive" /> : <CheckIcon className="size-3.5 text-success" />}
            <span className="flex-1 truncate">{TOOL_NAMES[t.name] ?? t.name}</span>
            {t.images ? <span className="text-[11px]">🖼 {t.images}</span> : null}
            {t.ms !== undefined && <span className="text-[11px] tabular-nums">{(t.ms / 1000).toFixed(1)} s</span>}
            <ChevronRightIcon className={cn("size-3.5 transition-transform", open === t.id && "rotate-90")} />
          </button>
          {open === t.id && (
            <div className="mx-1.5 mb-1 flex flex-col gap-1.5 rounded-lg bg-muted/60 p-2 font-mono text-[11.5px]">
              <pre className="whitespace-pre-wrap break-words">{JSON.stringify(t.args, null, 1)}</pre>
              {t.result && <pre className={cn("max-h-60 overflow-auto whitespace-pre-wrap break-words border-t pt-1.5", t.status === "error" && "text-destructive")}>{t.result}</pre>}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

function AssistantMessage({ item, last, status, mascot, onAnswer, disabled }: { item: Assistant; last: boolean; status?: string; mascot?: ReactNode; onAnswer: (t: string) => void; disabled: boolean }) {
  return (
    <div className="flex gap-2.5">
      <div className="w-8 shrink-0 pt-0.5">{mascot}</div>
      <div className="flex min-w-0 flex-1 flex-col gap-2">
        <Steps tools={item.tools} />
        {/* A question is shown alone: its card says everything (the text written beside it is not shown). */}
        {item.text && !item.ask && (
          <div className="text-[13.5px]">
            <Markdown>{item.text}</Markdown>
          </div>
        )}
        {item.ask && <Ask ask={item.ask} active={last && !disabled} onAnswer={onAnswer} />}
        {item.error && (
          <div className="flex items-start gap-2 rounded-xl border border-destructive/30 bg-destructive/5 px-3 py-2 text-[12.5px] text-destructive">
            <CircleAlertIcon className="mt-0.5 size-3.5 shrink-0" /> {item.error}
          </div>
        )}
        {item.running && status && (
          <div className="flex items-center gap-2 text-[12.5px] text-muted-foreground">
            <Loader2Icon className="size-3.5 animate-spin" /> {status}
          </div>
        )}
      </div>
    </div>
  );
}

function Ask({ ask, active, onAnswer }: { ask: NonNullable<Assistant["ask"]>; active: boolean; onAnswer: (t: string) => void }) {
  const [picked, setPicked] = useState<Record<number, string[]>>({});
  const complete = ask.questions.every((_, i) => picked[i]?.length);
  const answer = (p = picked) => onAnswer(ask.questions.map((q, i) => (ask.questions.length > 1 ? `${q.question} → ${(p[i] ?? []).join(", ")}` : (p[i] ?? []).join(", "))).join("\n"));
  return (
    <div className="flex flex-col gap-3 rounded-2xl border bg-muted/30 p-3.5">
      {ask.message && (
        <div className="text-[13.5px]">
          <Markdown>{ask.message}</Markdown>
        </div>
      )}
      {ask.questions.map((q, i) => (
        <div key={i} className="flex flex-col gap-2">
          <div className="text-[13px] font-semibold">{q.question}</div>
          <div className="flex flex-wrap gap-1.5">
            {q.options.map((o) => {
              const on = picked[i]?.includes(o);
              return (
                <button
                  key={o}
                  disabled={!active}
                  onClick={() => {
                    const next = { ...picked, [i]: q.multiple ? (on ? picked[i].filter((x) => x !== o) : [...(picked[i] ?? []), o]) : [o] };
                    setPicked(next);
                    // One single-choice question: answered at once.
                    if (ask.questions.length === 1 && !q.multiple) answer(next);
                  }}
                  className={cn("rounded-full border px-3 py-1.5 text-left text-[13px] transition-colors disabled:opacity-60", on ? "border-foreground bg-foreground text-background" : "bg-card hover:bg-muted")}
                >
                  {o}
                </button>
              );
            })}
          </div>
        </div>
      ))}
      {active && (ask.questions.length > 1 || ask.questions.some((q) => q.multiple)) && (
        <Button variant="dark" size="sm" className="self-start" disabled={!complete} onClick={() => answer()}>
          Send answers
        </Button>
      )}
    </div>
  );
}

function Composer({ chat, placeholder, disabled, accept }: { chat: Chat; placeholder?: string; disabled?: string; accept: string }) {
  const [text, setText] = useState("");
  const [files, setFiles] = useState<FileRef[]>([]);
  const [uploading, setUploading] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const area = useRef<HTMLTextAreaElement>(null);
  useEffect(() => {
    const el = area.current;
    if (el) (el.style.height = "auto"), (el.style.height = `${Math.min(200, el.scrollHeight)}px`);
  }, [text]);
  const add = async (list: FileList | File[]) => {
    setUploading(true);
    try {
      for (const f of Array.from(list)) {
        if (f.size > 20 * 1024 * 1024) throw new Error(`${f.name} is larger than 20 MB`);
        const ref = await keepFile(f, f.name || "picture.png");
        setFiles((x) => [...x, ref]);
      }
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setUploading(false);
    }
  };
  const send = () => {
    if (chat.running || disabled) return;
    void chat.send(text.trim(), files);
    setText("");
    setFiles([]);
  };
  return (
    <div className="border-t p-3">
      <div
        className="rounded-2xl border bg-card px-3 pt-2.5 pb-2 shadow-[0_1px_6px_rgba(0,0,0,0.04)] focus-within:border-foreground/30"
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          e.preventDefault();
          if (e.dataTransfer.files.length) void add(e.dataTransfer.files);
        }}
      >
        {files.length > 0 && (
          <div className="mb-2 flex flex-wrap gap-1.5">
            {files.map((f) => (
              <span key={f.id} className="flex items-center gap-1.5 rounded-lg border bg-muted py-0.5 pr-1 pl-1 text-[12px]">
                {f.type.startsWith("image/") ? <img src={fileUrl(f.id)} alt="" className="size-6 rounded object-cover" /> : null}
                <span className="max-w-32 truncate">{f.name}</span>
                <button onClick={() => setFiles((x) => x.filter((y) => y.id !== f.id))} aria-label="Remove">
                  <XIcon className="size-3" />
                </button>
              </span>
            ))}
          </div>
        )}
        <textarea
          ref={area}
          value={text}
          rows={1}
          disabled={!!disabled}
          onChange={(e) => setText(e.target.value)}
          onPaste={(e) => {
            const pasted = [...e.clipboardData.files];
            if (pasted.length) (e.preventDefault(), void add(pasted));
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) (e.preventDefault(), send());
          }}
          placeholder={disabled ?? placeholder ?? "Ask the assistant…"}
          className="max-h-[200px] w-full resize-none bg-transparent text-[13.5px] outline-none placeholder:text-muted-foreground/70"
        />
        <div className="mt-1 flex items-center gap-1">
          <input ref={input} type="file" accept={accept} multiple hidden onChange={(e) => e.target.files && void add(e.target.files)} />
          <Button variant="ghost" size="icon-sm" onClick={() => input.current?.click()} disabled={!!disabled || uploading} aria-label="Attach a file">
            {uploading ? <Loader2Icon className="animate-spin" /> : <PaperclipIcon />}
          </Button>
          <div className="flex-1" />
          {chat.running ? (
            <Button variant="dark" size="icon-sm" onClick={chat.stop} aria-label="Stop">
              <SquareIcon className="size-3 fill-current" />
            </Button>
          ) : (
            <Button variant="dark" size="icon-sm" onClick={send} disabled={!!disabled || (!text.trim() && !files.length)} aria-label="Send">
              <ArrowUpIcon />
            </Button>
          )}
        </div>
      </div>
    </div>
  );
}
