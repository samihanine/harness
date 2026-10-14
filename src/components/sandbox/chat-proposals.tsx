/**
 * Sandbox: three proposals for the AI chat, on the same fake conversation (a request, tool steps with
 * one failure, a markdown answer with a table, a follow-up), each in a frame as wide as the side panel.
 * Front end only: the composer and buttons do nothing.
 */
import { useState, type ReactNode } from "react";
import { ArrowUpIcon, CheckIcon, ChevronDownIcon, ChevronRightIcon, CopyIcon, RotateCcwIcon, SparklesIcon, ThumbsDownIcon, ThumbsUpIcon, WrenchIcon, XIcon } from "lucide-react";
import { Markdown } from "@/components/markdown";

type Step = { name: string; args: string; result?: string; error?: string; ms: number };
type Message = { role: "user"; text: string; time: string } | { role: "assistant"; steps: Step[]; text: string; time: string; running?: boolean };

export const CONVERSATION: Message[] = [
  { role: "user", text: "Show me only the urgent tasks and tell me who is overloaded", time: "10:42" },
  {
    role: "assistant",
    time: "10:42",
    steps: [
      { name: "read_rows", args: '{"all": true}', result: "25 rows read.", ms: 120 },
      { name: "set_view", args: '{"filters": [{"field": "priority", "op": "is", "values": ["Hight"]}]}', error: "priority has no value \"Hight\". Its values: Low, Medium, High, Critical.", ms: 40 },
      { name: "set_view", args: '{"filters": [{"field": "priority", "op": "is", "values": ["High", "Critical"]}]}', result: "View updated.", ms: 35 },
    ],
    text: "**10 urgent tasks** are now shown (High or Critical).\n\n| Owner | Urgent | Due this week |\n|---|---|---|\n| Bastien | 4 | 2 |\n| David | 3 | 1 |\n| Alice | 2 | 0 |\n| Chloé | 1 | 0 |\n\n**Bastien** looks overloaded: 4 urgent tasks, 2 of them due before Friday. Should I move one to Chloé?",
  },
  { role: "user", text: 'Yes, move "Fix the invoice export bug" to Chloé', time: "10:43" },
  {
    role: "assistant",
    time: "10:43",
    steps: [{ name: "update_rows", args: '{"rows": [{"id": "TASK_006", "values": {"owner": "Chloé"}}]}', result: "1 row(s) updated.", ms: 610 }],
    text: "Done: **TASK_006 · Fix the invoice export bug** is now assigned to **Chloé**. Bastien has 3 urgent tasks left.",
  },
  { role: "user", text: "And sort them by due date", time: "10:44" },
  { role: "assistant", time: "10:44", running: true, steps: [{ name: "set_view", args: '{"sorts": [{"field": "due"}]}', ms: 0 }], text: "" },
];

const SUGGESTIONS = ["Group by status", "Who has the most overdue tasks?", "Add a task"];

/** The three frames, side by side. */
export function ChatProposals() {
  const frames: { name: string; pitch: string; body: ReactNode }[] = [
    { name: "1 · Bubbles", pitch: "Messenger-like: your messages in solid primary bubbles on the right, the assistant in bordered white bubbles with an avatar; the tool steps folded into one line inside the answer; suggestions above the composer.", body: <Bubbles /> },
    { name: "2 · Timeline", pitch: "Document-like: full-width turns with names; the steps as a timeline (green / red dots, durations); the answer in a card with an accent border and actions (copy, retry, 👍 👎); the composer shows what the AI sees.", body: <Timeline /> },
    { name: "3 · Focus cards", pitch: "High contrast for reading: the answer in a card whose first sentence is the headline; the steps as small status pills; larger text; a live progress card while the AI works.", body: <Focus /> },
  ];
  return (
    <div className="flex min-h-0 flex-1 gap-6 overflow-x-auto bg-muted/40 p-6">
      {frames.map((f) => (
        <div key={f.name} className="flex w-[380px] shrink-0 flex-col gap-2">
          <div>
            <h2 className="font-semibold">{f.name}</h2>
            <p className="text-[12px] text-muted-foreground">{f.pitch}</p>
          </div>
          <div className="flex h-[720px] flex-col overflow-hidden rounded-xl border bg-card shadow-sm">{f.body}</div>
        </div>
      ))}
    </div>
  );
}

const Header = () => (
  <div className="flex h-10 shrink-0 items-center gap-1 border-b px-3 text-[12px]">
    <span className="min-w-0 flex-1 truncate">Urgent tasks and workload</span>
    <ChevronDownIcon className="size-3.5 text-muted-foreground" />
  </div>
);

/* --------------------------------- 1 · Bubbles -------------------------------- */

function Bubbles() {
  return (
    <>
      <Header />
      <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto bg-zinc-50 px-3 py-4">
        {CONVERSATION.map((m, i) =>
          m.role === "user" ? (
            <div key={i} className="flex flex-col items-end gap-1">
              <div className="max-w-[85%] rounded-2xl rounded-br-md bg-primary px-3.5 py-2 text-[13px] leading-relaxed text-primary-foreground shadow-sm">{m.text}</div>
              <span className="pr-1 text-[10px] text-muted-foreground">{m.time}</span>
            </div>
          ) : (
            <div key={i} className="flex items-start gap-2">
              <span className="mt-0.5 flex size-7 shrink-0 items-center justify-center rounded-full bg-gradient-to-br from-violet-500 to-blue-500 text-white">
                <SparklesIcon className="size-3.5" />
              </span>
              <div className="flex min-w-0 flex-col gap-1">
                <div className="rounded-2xl rounded-tl-md border bg-white px-3.5 py-2.5 text-[13px] shadow-sm">
                  <FoldedSteps steps={m.steps} running={m.running} />
                  {m.running ? <Dots /> : <Markdown>{m.text}</Markdown>}
                </div>
                <span className="pl-1 text-[10px] text-muted-foreground">{m.time}</span>
              </div>
            </div>
          ),
        )}
      </div>
      <div className="flex flex-col gap-2 border-t bg-card p-2.5">
        <div className="flex gap-1.5 overflow-x-auto">
          {SUGGESTIONS.map((s) => (
            <button key={s} type="button" className="shrink-0 rounded-full border bg-card px-2.5 py-1 text-[11px] text-muted-foreground hover:border-ring hover:text-foreground">
              {s}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-1 rounded-full border bg-zinc-50 py-1 pr-1 pl-4 focus-within:border-ring">
          <input className="flex-1 bg-transparent text-[13px] outline-none" placeholder="Ask to filter, fix or add rows…" />
          <button type="button" className="flex size-8 items-center justify-center rounded-full bg-primary text-primary-foreground">
            <ArrowUpIcon className="size-4" />
          </button>
        </div>
      </div>
    </>
  );
}

/** "3 actions · 1 retried", unfolded on click: what the assistant did, out of the way. */
function FoldedSteps({ steps, running }: { steps: Step[]; running?: boolean }) {
  const [open, setOpen] = useState(false);
  if (!steps.length) return null;
  const failed = steps.filter((s) => s.error).length;
  return (
    <div className="mb-2 border-b pb-2">
      <button type="button" onClick={() => setOpen(!open)} className="flex items-center gap-1.5 text-[11px] text-muted-foreground hover:text-foreground">
        <WrenchIcon className="size-3" />
        {running ? "Working…" : `${steps.length} action${steps.length > 1 ? "s" : ""}${failed ? ` · ${failed} retried` : ""}`}
        <ChevronRightIcon className={`size-3 transition-transform ${open ? "rotate-90" : ""}`} />
      </button>
      {open && (
        <ul className="mt-1.5 flex flex-col gap-1">
          {steps.map((s, i) => (
            <li key={i} className="flex items-start gap-1.5 text-[11px]">
              {s.error ? <XIcon className="mt-0.5 size-3 shrink-0 text-red-500" /> : <CheckIcon className="mt-0.5 size-3 shrink-0 text-emerald-600" />}
              <span className="min-w-0">
                <span className="font-mono">{s.name}</span>
                <span className="block text-muted-foreground">{s.error ?? s.result}</span>
              </span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

const Dots = () => (
  <span className="flex gap-1 py-1">
    {[0, 1, 2].map((i) => (
      <span key={i} className="size-1.5 animate-bounce rounded-full bg-muted-foreground" style={{ animationDelay: `${i * 120}ms` }} />
    ))}
  </span>
);

/* --------------------------------- 2 · Timeline -------------------------------- */

function Timeline() {
  return (
    <>
      <Header />
      <div className="flex min-h-0 flex-1 flex-col gap-5 overflow-y-auto px-4 py-4">
        {CONVERSATION.map((m, i) =>
          m.role === "user" ? (
            <div key={i} className="flex flex-col gap-1.5">
              <span className="flex items-center gap-2 text-[11px] font-semibold text-muted-foreground uppercase">
                <span className="flex size-5 items-center justify-center rounded-full bg-zinc-800 text-[10px] text-white normal-case">S</span> You · {m.time}
              </span>
              <p className="rounded-lg bg-muted px-3 py-2 text-[13px] leading-relaxed">{m.text}</p>
            </div>
          ) : (
            <div key={i} className="flex flex-col gap-2">
              <span className="flex items-center gap-2 text-[11px] font-semibold text-muted-foreground uppercase">
                <span className="flex size-5 items-center justify-center rounded-full bg-blue-600 text-white"><SparklesIcon className="size-3" /></span> Assistant · {m.time}
              </span>
              <ol className="ml-2.5 flex flex-col gap-2 border-l pl-4">
                {m.steps.map((s, k) => (
                  <li key={k} className="relative text-[11px]">
                    <span className={`absolute top-1 -left-[21px] size-2.5 rounded-full ring-2 ring-card ${m.running ? "animate-pulse bg-blue-500" : s.error ? "bg-red-500" : "bg-emerald-500"}`} />
                    <span className="flex items-center gap-1.5">
                      <span className="font-mono font-medium">{s.name}</span>
                      {!m.running && <span className="text-muted-foreground tabular-nums">{s.ms} ms</span>}
                    </span>
                    <span className={s.error ? "text-red-600" : "text-muted-foreground"}>{m.running ? "Running…" : s.error ?? s.result}</span>
                  </li>
                ))}
              </ol>
              {!m.running && (
                <div className="rounded-lg border border-l-4 border-l-blue-600 bg-card px-3.5 py-3 text-[13px] shadow-sm">
                  <Markdown>{m.text}</Markdown>
                  <div className="mt-2.5 flex gap-0.5 border-t pt-2 text-muted-foreground">
                    {[[<CopyIcon key="c" />, "Copy"], [<RotateCcwIcon key="r" />, "Retry"], [<ThumbsUpIcon key="u" />, "Good"], [<ThumbsDownIcon key="d" />, "Bad"]].map(([icon, label]) => (
                      <button key={label as string} type="button" title={label as string} className="icon-btn size-7 [&_svg]:size-3.5">{icon}</button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          ),
        )}
      </div>
      <div className="m-2.5 flex flex-col gap-1.5 rounded-lg border bg-card p-2 focus-within:border-ring">
        <textarea rows={2} className="resize-none bg-transparent px-1 text-[13px] outline-none" placeholder="Ask to filter, fix or add rows…" />
        <div className="flex items-center gap-1.5 text-[11px] text-muted-foreground">
          <span className="rounded bg-muted px-1.5 py-0.5">Sees: view "Urgent" · 10 rows</span>
          <span className="rounded bg-muted px-1.5 py-0.5">gpt-6-luna</span>
          <button type="button" className="ml-auto flex size-7 items-center justify-center rounded-md bg-primary text-primary-foreground">
            <ArrowUpIcon className="size-4" />
          </button>
        </div>
      </div>
    </>
  );
}

/* ------------------------------- 3 · Focus cards ------------------------------- */

function Focus() {
  return (
    <>
      <Header />
      <div className="flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto bg-zinc-100/70 px-3 py-4">
        <div className="flex items-center gap-2 text-[10px] font-medium text-muted-foreground uppercase">
          <span className="h-px flex-1 bg-border" /> Today <span className="h-px flex-1 bg-border" />
        </div>
        {CONVERSATION.map((m, i) => {
          if (m.role === "user")
            return (
              <div key={i} className="ml-10 shrink-0 self-end rounded-2xl bg-blue-600 px-4 py-2.5 text-[14px] leading-snug text-white shadow">
                {m.text}
              </div>
            );
          if (m.running)
            return (
              <div key={i} className="flex shrink-0 flex-col gap-2 rounded-xl border-2 border-dashed border-blue-300 bg-white p-3">
                <span className="flex items-center gap-2 text-[13px] font-medium text-blue-700">
                  <span className="size-2 animate-ping rounded-full bg-blue-500" /> Working on it…
                </span>
                {m.steps.map((s, k) => (
                  <span key={k} className="text-[12px] text-muted-foreground">→ Sorting the view by due date</span>
                ))}
              </div>
            );
          const [headline, ...rest] = m.text.split("\n\n");
          return (
            <div key={i} className="shrink-0 overflow-hidden rounded-xl border bg-white shadow-sm">
              <div className="flex flex-wrap items-center gap-1 border-b bg-zinc-50 px-3 py-1.5">
                <SparklesIcon className="mr-1 size-3.5 text-violet-600" />
                {m.steps.map((s, k) => (
                  <span key={k} title={s.error ?? s.result} className={`rounded-full px-2 py-0.5 font-mono text-[10px] ${s.error ? "bg-red-100 text-red-700 line-through" : "bg-emerald-100 text-emerald-800"}`}>
                    {s.error ? "✗" : "✓"} {s.name}
                  </span>
                ))}
                <span className="ml-auto text-[10px] text-muted-foreground">{m.time}</span>
              </div>
              <div className="flex flex-col gap-2 px-4 py-3">
                <div className="text-[15px] leading-snug font-medium [&_strong]:text-blue-700"><Markdown>{headline}</Markdown></div>
                {rest.length > 0 && <div className="text-[14px] leading-relaxed"><Markdown>{rest.join("\n\n")}</Markdown></div>}
              </div>
            </div>
          );
        })}
      </div>
      <div className="flex items-end gap-2 border-t bg-card p-3">
        <textarea rows={1} className="max-h-32 flex-1 resize-none rounded-xl border bg-zinc-50 px-3 py-2.5 text-[14px] outline-none focus:border-blue-500" placeholder="Ask anything about this table…" />
        <button type="button" className="flex size-10 items-center justify-center rounded-xl bg-blue-600 text-white shadow">
          <ArrowUpIcon className="size-5" />
        </button>
      </div>
    </>
  );
}
