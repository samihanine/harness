/** Settings in four tabs: General, AI, Power BI, Confluence. Secrets are written to the server, never read back. */
import { useEffect, useRef, useState, type ReactNode } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BarChart3Icon, BookOpenTextIcon, BotIcon, CheckIcon, DatabaseIcon, EyeIcon, EyeOffIcon, DownloadIcon, ExternalLinkIcon, FolderOpenIcon, Loader2Icon, LogOutIcon, PaletteIcon, PlugZapIcon, PlusIcon, RefreshCwIcon, RotateCcwIcon, Trash2Icon, UploadIcon } from "lucide-react";
import { toast } from "sonner";
import { SCREENS, type LibraryKind, type Provider, type SettingsPatch } from "@shared/types";
import { Workspace } from "@/components/workspace";
import { SCREEN_INFO } from "@/components/screens";
import { ErrorBox, Loading } from "@/components/common";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { Mascot, OUTFITS, OUTFIT_LABELS, STATES, STATE_LABELS, type MascotState, type Outfit } from "@/features/mascot";
import { recordGif } from "@/features/mascot/gif";
import { download, errorText, get, post } from "@/lib/api";
import { refreshProject, useProjectList } from "@/lib/projects";
import { libraryEdit, openReport, refreshLibrary, useHidden, useLibrary } from "@/lib/library";
import { reportStructure } from "@/lib/items";
import { AddDialog } from "@/components/pickers";
import { resetSettings, saveSettings, useSettings, type AppSettings } from "@/lib/settings";
import { exportAll, importAll, list } from "@/lib/store";
import { useTheme } from "@/lib/theme";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/settings")({ component: Settings, validateSearch: (q: Record<string, unknown>) => ({ tab: typeof q.tab === "string" ? q.tab : undefined }) });

type TabId = "general" | "ai" | "powerbi" | "confluence" | "reports" | "models";

function Field({ label, hint, children }: { label: string; hint?: ReactNode; children: ReactNode }) {
  return (
    <div className="grid gap-1.5 sm:grid-cols-[220px_1fr] sm:items-start sm:gap-6">
      <div className="pt-2">
        <div className="text-[13.5px] font-medium">{label}</div>
        {hint && <div className="text-[12px] text-muted-foreground">{hint}</div>}
      </div>
      <div>{children}</div>
    </div>
  );
}

const Card = ({ title, children, reset, aside }: { title: string; children: ReactNode; reset?: "general" | "ai" | "powerbi" | "confluence"; aside?: ReactNode }) => (
  <section className="flex flex-col gap-5 rounded-3xl border p-6">
    <div className="flex items-center gap-2">
      <h2 className="flex-1 text-[16px] font-semibold">{title}</h2>
      {aside}
      {reset && (
        <Button variant="ghost" size="sm" title="Back to the values of the .env file (and the app's defaults)" onClick={() => confirm("Reset this part to the .env values?") && void resetSettings(reset).then(() => toast.success("Reset to the .env values"), (e) => toast.error(errorText(e)))}>
          <RotateCcwIcon /> Reset to .env
        </Button>
      )}
    </div>
    {children}
  </section>
);

/** "Test the connection": a call that says what works, or why not. */
function TestButton({ path, body }: { path: string; body?: unknown }) {
  const [state, setState] = useState<{ busy?: boolean; ok?: boolean; text?: string }>({});
  return (
    <div className="flex flex-wrap items-center gap-3">
      <Button variant="outline" size="sm" disabled={state.busy} onClick={() => (setState({ busy: true }), post<{ text: string }>(path, body ?? {}).then((r) => setState({ ok: true, text: r.text }), (e) => setState({ ok: false, text: errorText(e) })))}>
        {state.busy ? <Loader2Icon className="animate-spin" /> : <PlugZapIcon />} Test the connection
      </Button>
      {state.text && <span className={cn("text-[13px]", state.ok ? "text-success" : "text-destructive")}>{state.ok ? "✓" : "✗"} {state.text}</span>}
    </div>
  );
}

/** Saves on change, with a toast. */
function useSave() {
  return async (patch: SettingsPatch) => {
    try {
      await saveSettings(patch);
      toast.success("Saved");
    } catch (e) {
      toast.error(errorText(e));
    }
  };
}

function SecretInput({ set, placeholder, onSave }: { set: boolean; placeholder?: string; onSave: (v: string) => void }) {
  const [v, setV] = useState("");
  return (
    <div className="flex gap-2">
      <input className="input" type="password" value={v} onChange={(e) => setV(e.target.value)} placeholder={set ? "•••••••• (saved — type to replace)" : (placeholder ?? "Not set")} autoComplete="off" />
      <Button variant="outline" size="sm" className="h-9" disabled={!v} onClick={() => (onSave(v), setV(""))}>Save</Button>
    </div>
  );
}

function TextSetting({ value, onSave, placeholder, mono }: { value: string; onSave: (v: string) => void; placeholder?: string; mono?: boolean }) {
  const [v, setV] = useState(value);
  useEffect(() => setV(value), [value]);
  return <input className={cn("input", mono && "font-mono text-[12.5px]")} value={v} placeholder={placeholder} onChange={(e) => setV(e.target.value)} onBlur={() => v !== value && onSave(v)} />;
}

function MascotGifs() {
  const [outfit, setOutfit] = useState<Outfit>("trainer");
  const [state, setState] = useState<MascotState>("idle");
  const [busy, setBusy] = useState(false);
  const box = useRef<HTMLDivElement>(null);
  const record = async () => {
    setBusy(true);
    try {
      download(await recordGif(() => box.current?.querySelector("svg") ?? null, { seconds: 3, fps: 20, size: 260 }), `mascot-${outfit}-${state}.gif`);
    } catch (e) {
      toast.error(errorText(e));
    } finally {
      setBusy(false);
    }
  };
  return (
    <div className="flex flex-wrap items-center gap-6">
      <div ref={box} className="flex size-40 items-center justify-center rounded-3xl bg-muted/60"><Mascot outfit={outfit} state={state} size={140} /></div>
      <div className="flex flex-col gap-3">
        <select className="input" value={outfit} onChange={(e) => setOutfit(e.target.value as Outfit)}>{(Object.keys(OUTFITS) as Outfit[]).map((o) => <option key={o} value={o}>{OUTFIT_LABELS[o]}</option>)}</select>
        <select className="input" value={state} onChange={(e) => setState(e.target.value as MascotState)}>{STATES.map((s) => <option key={s} value={s}>{STATE_LABELS[s]}</option>)}</select>
        <Button variant="outline" size="sm" onClick={record} disabled={busy}>{busy ? <Loader2Icon className="animate-spin" /> : <DownloadIcon />} Download as GIF</Button>
      </div>
    </div>
  );
}

function General({ s }: { s: AppSettings }) {
  const save = useSave();
  const { theme, setTheme } = useTheme();
  const file = useRef<HTMLInputElement>(null);
  return (
    <div className="flex flex-col gap-5">
      <Card title="Look" reset="general">
        <Field label="Primary color" hint="Buttons, highlights, charts.">
          <div className="flex items-center gap-3">
            <input type="color" value={s.general.primaryColor} onChange={(e) => void save({ general: { primaryColor: e.target.value } })} className="size-10 cursor-pointer rounded-xl border bg-transparent" />
            {["#b39251", "#d97757", "#363d4e", "#2563eb", "#0f766e", "#111111"].map((c) => (
              <button key={c} onClick={() => void save({ general: { primaryColor: c } })} className="flex size-8 items-center justify-center rounded-full ring-offset-2" style={{ background: c }} aria-label={c}>
                {s.general.primaryColor === c && <CheckIcon className="size-4 text-white" />}
              </button>
            ))}
          </div>
        </Field>
        <Field label="Theme">
          <div className="flex gap-2">{(["light", "dark", "system"] as const).map((t) => <Button key={t} variant={theme === t ? "dark" : "outline"} size="sm" onClick={() => setTheme(t)}>{t[0].toUpperCase() + t.slice(1)}</Button>)}</div>
        </Field>
        <Field label="Screens" hint="Hide the screens you do not use.">
          <div className="grid gap-2 sm:grid-cols-2">
            {SCREENS.map((id) => (
              <label key={id} className="flex items-center gap-3 rounded-xl border px-3 py-2 text-[13.5px]">
                <Switch checked={s.general.screens[id] !== false} onCheckedChange={(v) => void save({ general: { screens: { ...s.general.screens, [id]: v } } })} />
                {SCREEN_INFO[id].label}
              </label>
            ))}
          </div>
        </Field>
      </Card>
      <Card title="Mascot">
        <Field label="Show the mascot" hint="The hedgehog of the assistants.">
          <Switch checked={s.general.mascot} onCheckedChange={(v) => void save({ general: { mascot: v } })} />
        </Field>
        <Field label="Animations" hint="Each outfit and state, as a GIF."><MascotGifs /></Field>
      </Card>
      <Card title="Data of this browser">
        <Field label="Export / import" hint="Items, analyses, decks, conversations, context links.">
          <div className="flex gap-2">
            <Button variant="outline" size="sm" onClick={async () => download(await exportAll(), "hub-data.json")}><DownloadIcon /> Export</Button>
            <input ref={file} type="file" accept=".json" hidden onChange={async (e) => e.target.files?.[0] && (await importAll(e.target.files[0]), toast.success("Imported"))} />
            <Button variant="outline" size="sm" onClick={() => file.current?.click()}><UploadIcon /> Import</Button>
          </div>
        </Field>
      </Card>
    </div>
  );
}

const PROVIDERS: { id: Provider; label: string; hint: string }[] = [
  { id: "maia", label: "Maia (TrustLayer)", hint: "The group's gateway to validated models: client id and secret." },
  { id: "claude-code", label: "Claude Code", hint: "Claude Code run by the server (Agent SDK), through your gateway (Claude Maison) or a key." },
  { id: "anthropic", label: "Anthropic API", hint: "Claude with an Anthropic API key (tests)." },
];

function Ai({ s }: { s: AppSettings }) {
  const save = useSave();
  return (
    <div className="flex flex-col gap-5">
      <Card title="Model" reset="ai">
        <Field label="Provider">
          <div className="grid gap-2">
            {PROVIDERS.map((p) => (
              <button key={p.id} onClick={() => void save({ ai: { provider: p.id } })} className={cn("flex items-start gap-3 rounded-2xl border px-4 py-3 text-left", s.ai.provider === p.id && "border-foreground/40 bg-muted/50")}>
                <span className={cn("mt-0.5 flex size-4 items-center justify-center rounded-full border", s.ai.provider === p.id && "border-foreground bg-foreground")} />
                <span><span className="block text-[14px] font-medium">{p.label}</span><span className="text-[12.5px] text-muted-foreground">{p.hint}</span></span>
              </button>
            ))}
          </div>
        </Field>
        <Field label="Model">
          <div className="flex gap-2">
            <select className="input" value={s.models.includes(s.ai.model) ? s.ai.model : ""} onChange={(e) => e.target.value && void save({ ai: { model: e.target.value } })}>
              {!s.models.includes(s.ai.model) && <option value="">{s.ai.model} (custom)</option>}
              {s.models.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
            <TextSetting value={s.ai.model} onSave={(v) => void save({ ai: { model: v.trim() } })} mono />
          </div>
        </Field>
        <Field label="Effort" hint="Deeper thinking: better answers, slower.">
          <div className="flex gap-2">{(["low", "medium", "high"] as const).map((e) => <Button key={e} size="sm" variant={s.ai.effort === e ? "dark" : "outline"} onClick={() => void save({ ai: { effort: e } })}>{e}</Button>)}</div>
        </Field>
        <Field label="Check" hint="A short request to the chosen model.">
          <TestButton path="/settings/test-ai" />
        </Field>
        <Field label="Maximum steps" hint="Tool calls per message.">
          <input type="number" min={5} max={200} className="input w-32" defaultValue={s.ai.maxSteps} onBlur={(e) => void save({ ai: { maxSteps: Number(e.target.value) || 40 } })} />
        </Field>
      </Card>
      {s.ai.provider === "anthropic" && (
        <Card title="Anthropic API">
          <Field label="API key"><SecretInput set={s.ai.anthropic.keySet} placeholder="sk-ant-…" onSave={(v) => void save({ ai: { anthropicKey: v } })} /></Field>
        </Card>
      )}
      {s.ai.provider === "maia" && (
        <Card title="Maia (TrustLayer v2)">
          <Field label="Address"><TextSetting value={s.ai.maia.url} onSave={(v) => void save({ ai: { maia: { url: v } } })} mono /></Field>
          <Field label="Client id"><SecretInput set={s.ai.maia.clientIdSet} onSave={(v) => void save({ ai: { maia: { clientId: v } } })} /></Field>
          <Field label="Client secret"><SecretInput set={s.ai.maia.clientSecretSet} onSave={(v) => void save({ ai: { maia: { clientSecret: v } } })} /></Field>
          <Field label="Maison trigram" hint="For the TrustLayer's reporting."><TextSetting value={s.ai.maia.maison} onSave={(v) => void save({ ai: { maia: { maison: v } } })} /></Field>
          <Field label="Use case id"><TextSetting value={s.ai.maia.useCaseId} onSave={(v) => void save({ ai: { maia: { useCaseId: v } } })} /></Field>
          <Field label="Environment">
            <select className="input" value={s.ai.maia.environment} onChange={(e) => void save({ ai: { maia: { environment: e.target.value } } })}>{["PRD", "STA", "PPD", "QA", "DEV"].map((x) => <option key={x}>{x}</option>)}</select>
          </Field>
        </Card>
      )}
      {s.ai.provider === "claude-code" && (
        <Card title="Claude Code">
          <Field label="Gateway address" hint="ANTHROPIC_BASE_URL of Claude Maison (empty: Anthropic)."><TextSetting value={s.ai.claudeCode.baseUrl} onSave={(v) => void save({ ai: { claudeCode: { baseUrl: v } } })} mono /></Field>
          <Field label="Headers" hint="ANTHROPIC_CUSTOM_HEADERS, one per line."><TextSetting value={s.ai.claudeCode.headers} onSave={(v) => void save({ ai: { claudeCode: { headers: v } } })} mono /></Field>
          <Field label="Key script" hint="Script giving the token (e.g. ~/.claude-maison/okta-auth.sh)."><TextSetting value={s.ai.claudeCode.keyHelper} onSave={(v) => void save({ ai: { claudeCode: { keyHelper: v } } })} mono /></Field>
          <Field label="Or an API key"><SecretInput set={s.ai.claudeCode.keySet} onSave={(v) => void save({ ai: { claudeCode: { key: v } } })} /></Field>
        </Card>
      )}
      <Card title="Conversations">
        <Field label="Export" hint="Every conversation of this browser, with the steps of the assistant.">
          <Button variant="outline" size="sm" onClick={async () => download(new Blob([JSON.stringify(await list("conversation"), null, 1)], { type: "application/json" }), "conversations.json")}><DownloadIcon /> Export conversations</Button>
        </Field>
      </Card>
    </div>
  );
}

function PowerBi({ s }: { s: AppSettings }) {
  const save = useSave();
  const [auth, setAuth] = useState<{ signedIn: boolean; account?: string; signIn?: { userCode: string; verificationUri: string }; error?: string }>();
  const [busy, setBusy] = useState(false);
  const refresh = () => get("/auth").then(setAuth, (e) => setAuth({ signedIn: false, error: errorText(e) }));
  useEffect(() => void refresh(), []);
  useEffect(() => {
    if (!auth?.signIn) return;
    const t = setInterval(refresh, 3000);
    return () => clearInterval(t);
  }, [auth?.signIn]);
  const p = s.powerbi;
  return (
    <div className="flex flex-col gap-5">
      <Card title="Microsoft sign-in">
        {!auth ? (
          <Loading />
        ) : auth.signedIn ? (
          <Field label="Signed in" hint="Power BI, Fabric, OneDrive (previews of PowerPoint files).">
            <div className="flex items-center gap-3">
              <span className="text-[14px]">{auth.account}</span>
              <Button variant="outline" size="sm" onClick={async () => (await post("/auth/sign-out", {}), refresh())}><LogOutIcon /> Sign out</Button>
            </div>
          </Field>
        ) : auth.signIn ? (
          <div className="flex flex-col gap-3 rounded-2xl bg-muted/50 p-5">
            <div className="text-[14px]">Open <a className="font-medium text-primary underline" href={auth.signIn.verificationUri} target="_blank" rel="noreferrer">{auth.signIn.verificationUri} <ExternalLinkIcon className="inline size-3.5" /></a> and enter the code:</div>
            <div className="font-mono text-[28px] font-semibold tracking-[0.2em]">{auth.signIn.userCode}</div>
            <div className="flex items-center gap-2 text-[12.5px] text-muted-foreground"><Loader2Icon className="size-3.5 animate-spin" /> Waiting for the sign-in…</div>
          </div>
        ) : (
          <Field label="Not signed in" hint="One sign-in with your Microsoft account: no app registration.">
            <Button variant="dark" size="sm" disabled={busy} onClick={async () => { setBusy(true); try { await post("/auth/sign-in", {}); await refresh(); } catch (e) { toast.error(errorText(e)); } finally { setBusy(false); } }}>Sign in to Microsoft</Button>
            {auth.error && <ErrorBox className="mt-3">{auth.error}</ErrorBox>}
          </Field>
        )}
      </Card>
      <Card title="Connection">
        <Field label="Check" hint="Lists your reports and models (workspaces and apps)."><TestButton path="/powerbi/test" /></Field>
      </Card>
      <Card title="Performance limits" reset="powerbi">
        <div className="text-[12.5px] text-muted-foreground">Every DAX query of the app and the assistants follows these limits, so the semantic models stay fast for everyone.</div>
        <Field label="Rows per query"><input type="number" className="input w-40" defaultValue={p.maxRows} onBlur={(e) => void save({ powerbi: { maxRows: Number(e.target.value) || 2000 } })} /></Field>
        <Field label="Time limit (s)"><input type="number" className="input w-40" defaultValue={p.timeoutMs / 1000} onBlur={(e) => void save({ powerbi: { timeoutMs: (Number(e.target.value) || 20) * 1000 } })} /></Field>
        <Field label="Queries at once per model"><input type="number" className="input w-40" defaultValue={p.concurrency} onBlur={(e) => void save({ powerbi: { concurrency: Number(e.target.value) || 2 } })} /></Field>
        <Field label="Cache (s)" hint="The same query again within this time: answered from the cache."><input type="number" className="input w-40" defaultValue={p.cacheMs / 1000} onBlur={(e) => void save({ powerbi: { cacheMs: (Number(e.target.value) || 0) * 1000 } })} /></Field>
      </Card>
    </div>
  );
}

/** The projects found in the projects folder, with their kept pages: refresh, open in Confluence. */
function Projects() {
  const { projects, error } = useProjectList();
  const [busy, setBusy] = useState<string>();
  const refresh = (id: string) => (setBusy(id), refreshProject(id).catch((e) => toast.error(errorText(e))).finally(() => setBusy(undefined)));
  return (
    <Card title="Projects" aside={<Button variant="outline" size="sm" disabled={!!busy} onClick={() => refresh("*")}>{busy === "*" ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Read the list again</Button>}>
      <div className="text-[12.5px] text-muted-foreground">Each sub-folder of the projects folder is a project (its name). The pages of the projects chosen in the chat (top of the assistant) are given whole to it at the start of each conversation. They are read once and kept: "Refresh" after they change.</div>
      {error && <ErrorBox>{error}</ErrorBox>}
      {!projects && !error && <Loading />}
      {projects?.map((p) => (
        <div key={p.id} className="flex flex-wrap items-center gap-3 rounded-2xl border px-4 py-2.5">
          <FolderOpenIcon className="size-4 text-muted-foreground" />
          <span className="w-56 truncate text-[14px] font-medium">{p.name}</span>
          <span className="flex-1 text-[12.5px] text-muted-foreground">{p.error ? <span className="text-destructive">{p.error}</span> : p.pages === undefined ? "not read yet (read when first used)" : `${p.pages} pages · read ${new Date(p.at!).toLocaleString()}`}</span>
          <Button variant="ghost" size="sm" disabled={!!busy} onClick={() => refresh(p.id)}>{busy === p.id ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh</Button>
          {p.url && <Button variant="ghost" size="sm" asChild><a href={p.url} target="_blank" rel="noreferrer"><ExternalLinkIcon /> Open</a></Button>}
        </div>
      ))}
      {projects && !projects.length && <div className="text-[13px] text-muted-foreground">No project: set the projects folder above (one sub-folder per project).</div>}
    </Card>
  );
}

function ConfluenceSettings({ s }: { s: AppSettings }) {
  const save = useSave();
  const c = s.confluence;
  const folder = (label: string, key: "projectsFolderId" | "reportsFolderId" | "datasetsFolderId", env: string, hint: string) => (
    <Field label={label} hint={`${hint} (${env}).`}>
      <div className="flex items-center gap-2">
        <div className="flex-1"><TextSetting value={c[key]} onSave={(v) => void save({ confluence: { [key]: v.trim() } })} mono placeholder="Folder id or link" /></div>
        <TestButton path="/confluence/test" body={{ folderId: c[key] }} />
      </div>
    </Field>
  );
  return (
    <div className="flex flex-col gap-5">
      <Card title="Confluence" reset="confluence">
        <Field label="Address" hint="e.g. https://<site>.atlassian.net/wiki"><TextSetting value={c.url} onSave={(v) => void save({ confluence: { url: v.trim() } })} mono /></Field>
        <Field label="Email" hint="Your Atlassian account (Cloud). Empty: the token is a Data Center personal token."><TextSetting value={c.email} onSave={(v) => void save({ confluence: { email: v.trim() } })} /></Field>
        <Field label="API token"><SecretInput set={c.keySet} onSave={(v) => void save({ confluence: { key: v } })} /></Field>
        {folder("Projects folder", "projectsFolderId", "CONFLUENCE_PROJECTS_FOLDER_ID", "One sub-folder per project, named as the project")}
        {folder("Reports folder", "reportsFolderId", "CONFLUENCE_REPORT_FOLDER_ID", "The team's reports: one page each (its name, link and description)")}
        {folder("Models folder", "datasetsFolderId", "CONFLUENCE_DATASET_FOLDER_ID", "The team's semantic models: one page each")}
      </Card>
      <Projects />
    </div>
  );
}

/** The catalog of reports or models (Confluence pages): add, rename, describe, refresh, hide for me, remove. */
function LibrarySettings({ kind }: { kind: LibraryKind }) {
  const { entries, error } = useLibrary(kind);
  const [hidden, setHidden] = useHidden(kind);
  const [adding, setAdding] = useState(false);
  const [busy, setBusy] = useState<string>();
  const run = (key: string, f: () => Promise<unknown>, done?: string) => (setBusy(key), f().then(() => done && toast.success(done), (e) => toast.error(errorText(e))).finally(() => setBusy(undefined)));
  const what = kind === "report" ? "reports" : "semantic models";
  return (
    <Card title={kind === "report" ? "Reports" : "Semantic models"} aside={<div className="flex gap-2"><Button variant="outline" size="sm" disabled={!!busy} onClick={() => run("list", () => refreshLibrary(kind))}>{busy === "list" ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Read again</Button><Button variant="dark" size="sm" onClick={() => setAdding(true)}><PlusIcon /> Add</Button></div>}>
      <div className="text-[12.5px] text-muted-foreground">The team's {what}, one Confluence page each: its title is the name, its content the Power BI link and a description (what it is for — the assistants use it to choose). Hiding one hides it for you only.</div>
      {error && <ErrorBox>{error}</ErrorBox>}
      {!entries && !error && <Loading />}
      {entries?.map((e) => {
        const off = !!e.id && hidden.includes(e.id);
        return (
          <div key={e.pageId} className={cn("flex flex-col gap-2 rounded-2xl border p-3", off && "opacity-60")}>
            <div className="flex items-center gap-2">
              <input className="input h-8 flex-1 font-medium" defaultValue={e.name} onBlur={(x) => x.target.value.trim() && x.target.value.trim() !== e.name && run(e.pageId, () => libraryEdit.change(kind, e.pageId, { name: x.target.value.trim() }), "Renamed")} />
              {kind === "report" && e.id && <Button variant="ghost" size="sm" disabled={!!busy} title="Read its structure again (kept in this browser)" onClick={() => run(e.pageId, async () => reportStructure(await openReport(e, true), true), "Structure read again")}>{busy === e.pageId ? <Loader2Icon className="animate-spin" /> : <RefreshCwIcon />} Refresh</Button>}
              {e.id && <Button variant="ghost" size="icon-sm" title={off ? "Show it in my lists" : "Hide it from my lists"} onClick={() => setHidden(off ? hidden.filter((x) => x !== e.id) : [...hidden, e.id!])}>{off ? <EyeIcon /> : <EyeOffIcon />}</Button>}
              {e.url && <Button variant="ghost" size="icon-sm" asChild title="Open in Power BI"><a href={e.url} target="_blank" rel="noreferrer"><ExternalLinkIcon /></a></Button>}
              {e.pageUrl && <Button variant="ghost" size="icon-sm" asChild title="Open the Confluence page"><a href={e.pageUrl} target="_blank" rel="noreferrer"><BookOpenTextIcon /></a></Button>}
              <Button variant="ghost" size="icon-sm" title="Remove from the catalog" disabled={!!busy} onClick={() => confirm(`Remove "${e.name}" from the catalog (its Confluence page goes to the trash)?`) && run(e.pageId, () => libraryEdit.remove(kind, e.pageId), "Removed")}><Trash2Icon /></Button>
            </div>
            <textarea className="input h-auto min-h-14 py-1.5 text-[12.5px]" placeholder="What it is for, who uses it, what it shows…" defaultValue={e.description} onBlur={(x) => x.target.value.trim() !== e.description && run(e.pageId, () => libraryEdit.change(kind, e.pageId, { description: x.target.value }), "Description saved")} />
            {e.error && <div className="text-[12px] text-destructive">{e.error}</div>}
          </div>
        );
      })}
      {entries && !entries.length && <div className="text-[13px] text-muted-foreground">No {what} yet: add one.</div>}
      <AddDialog kind={kind} open={adding} onOpenChange={setAdding} onAdd={(link) => libraryEdit.add(kind, link)} />
    </Card>
  );
}

function Settings() {
  const s = useSettings();
  // The tab in the address (?tab=reports: "Manage the list" of the screens' pickers).
  const search = Route.useSearch();
  const navigate = Route.useNavigate();
  const tab = (search.tab ?? "general") as TabId;
  const setTab = (t: TabId) => void navigate({ search: { tab: t }, replace: true });
  const catalog = !!s?.confluence.catalogEdit;
  const tabs = [
    { id: "general" as const, label: "General", icon: <PaletteIcon /> },
    { id: "ai" as const, label: "AI", icon: <BotIcon /> },
    { id: "powerbi" as const, label: "Power BI", icon: <span className="text-[11px] font-bold">BI</span> },
    { id: "confluence" as const, label: "Confluence", icon: <BookOpenTextIcon /> },
    ...(catalog ? [{ id: "reports" as const, label: "Reports", icon: <BarChart3Icon /> }, { id: "models" as const, label: "Models", icon: <DatabaseIcon /> }] : []),
  ];
  return (
    <Workspace screen="settings" title="Settings" tabs={tabs} tab={tab} onTab={setTab}>
      <div className="mx-auto max-w-4xl p-6">{!s ? <Loading /> : tab === "general" ? <General s={s} /> : tab === "ai" ? <Ai s={s} /> : tab === "powerbi" ? <PowerBi s={s} /> : tab === "reports" && catalog ? <LibrarySettings kind="report" /> : tab === "models" && catalog ? <LibrarySettings kind="dataset" /> : <ConfluenceSettings s={s} />}</div>
    </Workspace>
  );
}
