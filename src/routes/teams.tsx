/**
 * Teams: the assistant of one of the user's Teams chats (a group or a one-to-one chat). Its name, description
 * and instructions, what it may use (projects, reports, models), when it answers (whatever concerns it, or when
 * named), at most how often; on / off; what it did. The mascot's GIFs it sends are made here, once.
 */
import { useEffect, useRef, useState } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BotIcon, CheckIcon, Loader2Icon, MessagesSquareIcon, RefreshCwIcon, SparklesIcon } from "lucide-react";
import { toast } from "sonner";
import { Workspace } from "@/components/workspace";
import { ErrorBox, Loading } from "@/components/common";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { Mascot, type MascotState } from "@/features/mascot";
import { recordGif } from "@/features/mascot/gif";
import { errorText, get, post } from "@/lib/api";
import { useLibrary } from "@/lib/library";
import { useProjectList } from "@/lib/projects";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/teams")({ component: Teams });

type Config = { enabled: boolean; chatId?: string; chatName?: string; name: string; description: string; instructions: string; projects: string[]; reportIds: string[]; datasetIds: string[]; maxPerHour: number; mode: "relevant" | "mentions" };
type Log = { at: string; from: string; text: string; reply?: string; gif?: string; skipped?: string; error?: string };
type Status = { config: Config; log: Log[]; status?: string; gifs: Record<string, boolean> };
type Chat = { id: string; name: string; type: string; last?: string };

function Card({ title, children, aside }: { title: string; children: React.ReactNode; aside?: React.ReactNode }) {
  return (
    <div className="flex flex-col gap-3 rounded-3xl border bg-card p-5">
      <div className="flex items-center gap-2">
        <div className="flex-1 text-[15px] font-semibold">{title}</div>
        {aside}
      </div>
      {children}
    </div>
  );
}

/** Checkable names (projects, reports, models). */
function Checks({ items, ids, onChange, empty }: { items: { id: string; name: string }[]; ids: string[]; onChange: (ids: string[]) => void; empty: string }) {
  if (!items.length) return <div className="text-[13px] text-muted-foreground">{empty}</div>;
  return (
    <div className="flex flex-wrap gap-1.5">
      {items.map((x) => {
        const on = ids.includes(x.id);
        return (
          <button key={x.id} onClick={() => onChange(on ? ids.filter((i) => i !== x.id) : [...ids, x.id])} className={cn("flex items-center gap-1.5 rounded-full border px-3 py-1 text-[13px]", on ? "border-foreground bg-foreground text-background" : "hover:bg-muted")}>
            {on && <CheckIcon className="size-3.5" />} {x.name}
          </button>
        );
      })}
    </div>
  );
}

const MOODS: MascotState[] = ["happy", "sad", "angry", "surprised", "party", "thinking"];

/** The mascot's GIFs for Teams: drawn here (the mascot is the app's), sent to the server once. */
function Gifs({ made, onMade }: { made: Record<string, boolean>; onMade: () => void }) {
  const box = useRef<HTMLDivElement>(null);
  const [busy, setBusy] = useState<string>();
  const make = async (moods: MascotState[]) => {
    for (const m of moods) {
      setBusy(m);
      const gif = await recordGif(() => box.current?.querySelector(`[data-mood="${m}"] svg`) ?? null, { seconds: 2.4, fps: 15, size: 200 });
      await fetch(`/api/teams/gif/${m}`, { method: "POST", body: gif });
    }
    setBusy(undefined);
    onMade();
  };
  useEffect(() => {
    const missing = MOODS.filter((m) => made[m] === false);
    if (missing.length) void make(missing);
  }, [JSON.stringify(made)]); // eslint-disable-line react-hooks/exhaustive-deps
  return (
    <Card title="GIFs" aside={<Button variant="ghost" size="sm" disabled={!!busy} onClick={() => void make(MOODS)}>{busy ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Make them again</Button>}>
      <div className="text-[12.5px] text-muted-foreground">The assistant sends one with a funny answer to a remark outside its job, or when someone has their answer.</div>
      <div ref={box} className="flex flex-wrap gap-3">
        {MOODS.map((m) => (
          <div key={m} data-mood={m} className="flex flex-col items-center gap-1 rounded-2xl border bg-white p-2">
            <Mascot outfit={m === "party" ? "party" : "diver"} state={m} size={90} />
            <span className="text-[11.5px] text-muted-foreground">{busy === m ? "recording…" : made[m] ? m : `${m} (not made)`}</span>
          </div>
        ))}
      </div>
    </Card>
  );
}

function Teams() {
  const [s, setS] = useState<Status>();
  const [chats, setChats] = useState<Chat[]>();
  const [chatError, setChatError] = useState<string>();
  const { projects } = useProjectList();
  const { entries: reports } = useLibrary("report");
  const { entries: models } = useLibrary("dataset");
  const load = () => get<Status>("/teams").then(setS, (e) => toast.error(errorText(e)));
  useEffect(() => {
    void load();
    get<Chat[]>("/teams/chats").then(setChats, (e) => setChatError(errorText(e)));
    const t = setInterval(() => void load(), 5000);
    return () => clearInterval(t);
  }, []);
  const save = (patch: Partial<Config>) => post<Status>("/teams", patch).then(setS, (e) => toast.error(errorText(e)));
  if (!s) return <Workspace screen="teams" title="Teams"><Loading /></Workspace>;
  const c = s.config;
  return (
    <Workspace screen="teams" title="Teams" item={<span className="text-[15px] text-muted-foreground">{c.chatName ?? "No chat chosen"}</span>} actions={<div className="flex items-center gap-2 text-[13px]"><span className={cn("size-2 rounded-full", c.enabled && c.chatId ? "bg-success" : "bg-muted-foreground/40")} />{c.enabled ? (c.chatId ? "On" : "Choose a chat") : "Off"}<Switch checked={c.enabled} onCheckedChange={(enabled) => void save({ enabled })} /></div>}>
      <div className="mx-auto flex max-w-4xl flex-col gap-5 p-6">
        <Card title="The chat">
          <div className="text-[12.5px] text-muted-foreground">One of your Teams chats (a group or a one-to-one chat). The assistant reads every new message — it is checked every few seconds while the app runs — and answers as you, signed with its name. It never answers its own messages.</div>
          {chatError && <ErrorBox>{chatError.includes("license") ? "Teams is not available for this Microsoft account (no Teams license)." : chatError}</ErrorBox>}
          {!chats && !chatError && <div className="flex items-center gap-2 text-[13px] text-muted-foreground"><Loader2Icon className="size-4 animate-spin" /> Reading your chats…</div>}
          {chats && (
            <select className="input" value={c.chatId ?? ""} onChange={(e) => { const ch = chats.find((x) => x.id === e.target.value); void save({ chatId: ch?.id, chatName: ch?.name }); }}>
              <option value="">Choose a chat…</option>
              {chats.map((x) => <option key={x.id} value={x.id}>{x.name} {x.type === "oneOnOne" ? "(one-to-one)" : ""}</option>)}
            </select>
          )}
          {s.status?.startsWith("error") && <ErrorBox>{s.status}</ErrorBox>}
        </Card>
        <Card title="The assistant" aside={<BotIcon className="size-4 text-muted-foreground" />}>
          <label className="label">Name (its signature, and how people call it)</label>
          <input className="input" defaultValue={c.name} onBlur={(e) => e.target.value.trim() && e.target.value !== c.name && void save({ name: e.target.value.trim() })} />
          <label className="label">Description (who it is: its job, its character — it answers off-topic remarks in character)</label>
          <textarea className="input h-auto min-h-20 py-2" defaultValue={c.description} onBlur={(e) => e.target.value !== c.description && void save({ description: e.target.value })} />
          <label className="label">Instructions</label>
          <textarea className="input h-auto min-h-24 py-2" placeholder="e.g. Answer in French. Revenue is always excluding tax. Never give figures by person." defaultValue={c.instructions} onBlur={(e) => e.target.value !== c.instructions && void save({ instructions: e.target.value })} />
          <div className="flex flex-wrap items-center gap-4">
            <label className="flex items-center gap-2 text-[13px]">
              Answers
              <select className="input w-64" value={c.mode} onChange={(e) => void save({ mode: e.target.value as Config["mode"] })}>
                <option value="relevant">whatever concerns it</option>
                <option value="mentions">only when its name is written</option>
              </select>
            </label>
            <label className="flex items-center gap-2 text-[13px]">
              At most
              <input type="number" className="input w-20" defaultValue={c.maxPerHour} onBlur={(e) => void save({ maxPerHour: Math.max(1, Number(e.target.value) || 20) })} />
              answers an hour
            </label>
          </div>
        </Card>
        <Card title="What it knows" aside={<SparklesIcon className="size-4 text-muted-foreground" />}>
          <div className="label">Projects (their Confluence pages are its context)</div>
          <Checks items={(projects ?? []).map((p) => ({ id: p.id, name: p.name }))} ids={c.projects} onChange={(projects) => void save({ projects })} empty="No project (Settings › Confluence)." />
          <div className="label">Reports</div>
          <Checks items={(reports ?? []).filter((e) => e.id).map((e) => ({ id: e.id!, name: e.name }))} ids={c.reportIds} onChange={(reportIds) => void save({ reportIds })} empty="No report in the catalog." />
          <div className="label">Semantic models (it queries them for the figures)</div>
          <Checks items={(models ?? []).filter((e) => e.id).map((e) => ({ id: e.id!, name: e.name }))} ids={c.datasetIds} onChange={(datasetIds) => void save({ datasetIds })} empty="No model in the catalog." />
        </Card>
        <Gifs made={s.gifs} onMade={() => void load()} />
        <Card title="What it did" aside={<MessagesSquareIcon className="size-4 text-muted-foreground" />}>
          {!s.log.length && <div className="text-[13px] text-muted-foreground">Nothing yet.</div>}
          {s.log.map((l, i) => (
            <div key={i} className="flex flex-col gap-1 rounded-2xl border px-4 py-2.5 text-[13px]">
              <div className="flex gap-2 text-[12px] text-muted-foreground"><span>{new Date(l.at).toLocaleString()}</span><span className="font-medium text-foreground">{l.from}</span></div>
              <div className="line-clamp-3">{l.text}</div>
              {l.reply && <div className="rounded-xl bg-muted/60 px-3 py-1.5 whitespace-pre-wrap">🤖 {l.reply}{l.gif ? ` [GIF ${l.gif}]` : ""}</div>}
              {l.skipped && <div className="text-[12px] text-muted-foreground">No answer: {l.skipped}</div>}
              {l.error && <div className="text-[12px] text-destructive">{l.error}</div>}
            </div>
          ))}
        </Card>
      </div>
    </Workspace>
  );
}
