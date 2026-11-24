/** Settings in four tabs: General, AI, Power BI, Confluence. Secrets are written to the server, never read back. */
import { useEffect, useRef, useState, type ReactNode } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { BookOpenTextIcon, BotIcon, CheckIcon, DownloadIcon, ExternalLinkIcon, Loader2Icon, LogOutIcon, PaletteIcon, UploadIcon } from "lucide-react";
import { toast } from "sonner";
import { SCREENS, type Provider, type SettingsPatch } from "@shared/types";
import { Workspace } from "@/components/workspace";
import { SCREEN_INFO } from "@/components/screens";
import { ErrorBox, Loading } from "@/components/common";
import { Button } from "@/components/ui/button";
import { Switch } from "@/components/ui/switch";
import { Mascot, OUTFITS, OUTFIT_LABELS, STATES, STATE_LABELS, type MascotState, type Outfit } from "@/features/mascot";
import { recordGif } from "@/features/mascot/gif";
import { download, errorText, get, post } from "@/lib/api";
import { saveSettings, useSettings, type AppSettings } from "@/lib/settings";
import { exportAll, importAll, list } from "@/lib/store";
import { useTheme } from "@/lib/theme";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/settings")({ component: Settings });

type TabId = "general" | "ai" | "powerbi" | "confluence";

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

const Card = ({ title, children }: { title: string; children: ReactNode }) => (
  <section className="flex flex-col gap-5 rounded-3xl border p-6">
    <h2 className="text-[16px] font-semibold">{title}</h2>
    {children}
  </section>
);

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
      <Card title="Look">
        <Field label="Primary color" hint="Buttons, highlights, charts.">
          <div className="flex items-center gap-3">
            <input type="color" value={s.general.primaryColor} onChange={(e) => void save({ general: { primaryColor: e.target.value } })} className="size-10 cursor-pointer rounded-xl border bg-transparent" />
            {["#d97757", "#b39251", "#2563eb", "#0f766e", "#7c3aed", "#111111"].map((c) => (
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
      <Card title="Model">
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
      <Card title="Performance limits">
        <div className="text-[12.5px] text-muted-foreground">Every DAX query of the app and the assistants follows these limits, so the semantic models stay fast for everyone.</div>
        <Field label="Rows per query"><input type="number" className="input w-40" defaultValue={p.maxRows} onBlur={(e) => void save({ powerbi: { maxRows: Number(e.target.value) || 2000 } })} /></Field>
        <Field label="Time limit (s)"><input type="number" className="input w-40" defaultValue={p.timeoutMs / 1000} onBlur={(e) => void save({ powerbi: { timeoutMs: (Number(e.target.value) || 20) * 1000 } })} /></Field>
        <Field label="Queries at once per model"><input type="number" className="input w-40" defaultValue={p.concurrency} onBlur={(e) => void save({ powerbi: { concurrency: Number(e.target.value) || 2 } })} /></Field>
        <Field label="Cache (s)" hint="The same query again within this time: answered from the cache."><input type="number" className="input w-40" defaultValue={p.cacheMs / 1000} onBlur={(e) => void save({ powerbi: { cacheMs: (Number(e.target.value) || 0) * 1000 } })} /></Field>
      </Card>
    </div>
  );
}

function ConfluenceSettings({ s }: { s: AppSettings }) {
  const save = useSave();
  const c = s.confluence;
  const [test, setTest] = useState<string>();
  return (
    <Card title="Confluence">
      <Field label="Address" hint="e.g. https://<site>.atlassian.net/wiki"><TextSetting value={c.url} onSave={(v) => void save({ confluence: { url: v.trim() } })} mono /></Field>
      <Field label="Email" hint="Your Atlassian account (Cloud). Empty: the token is a Data Center personal token."><TextSetting value={c.email} onSave={(v) => void save({ confluence: { email: v.trim() } })} /></Field>
      <Field label="API token"><SecretInput set={c.keySet} onSave={(v) => void save({ confluence: { key: v } })} /></Field>
      <Field label="Folder (root)" hint="Id of the folder or page whose pages the app uses."><TextSetting value={c.folderId} onSave={(v) => void save({ confluence: { folderId: v.trim() } })} mono /></Field>
      <Field label="Check">
        <div className="flex items-center gap-3">
          <Button variant="outline" size="sm" onClick={() => get<unknown[]>("/confluence/pages").then((p) => setTest(`✓ ${p.length} pages found`), (e) => setTest(`✗ ${errorText(e)}`))}>Test the connection</Button>
          {test && <span className={cn("text-[13px]", test.startsWith("✓") ? "text-success" : "text-destructive")}>{test}</span>}
        </div>
      </Field>
    </Card>
  );
}

function Settings() {
  const s = useSettings();
  const [tab, setTab] = useState<TabId>("general");
  return (
    <Workspace screen="reports" title="Settings" tabs={[{ id: "general", label: "General", icon: <PaletteIcon /> }, { id: "ai", label: "AI", icon: <BotIcon /> }, { id: "powerbi", label: "Power BI", icon: <span className="text-[11px] font-bold">BI</span> }, { id: "confluence", label: "Confluence", icon: <BookOpenTextIcon /> }]} tab={tab} onTab={setTab}>
      <div className="mx-auto max-w-4xl p-6">{!s ? <Loading /> : tab === "general" ? <General s={s} /> : tab === "ai" ? <Ai s={s} /> : tab === "powerbi" ? <PowerBi s={s} /> : <ConfluenceSettings s={s} />}</div>
    </Workspace>
  );
}
