import { useEffect, useRef, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { BookOpenIcon, FileTextIcon, GaugeIcon, UploadIcon, BarChart3Icon, DownloadIcon, LogOutIcon, MonitorIcon, MoonIcon, PresentationIcon, RotateCcwIcon, SaveIcon, SearchIcon, Settings2Icon, SettingsIcon, SmileIcon, SparklesIcon, SunIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { Section, Workspace } from "@/components/workspace";
import { ConversationsDialog } from "@/components/conversations";
import { Mascot, OUTFIT_LABELS, STATES, STATE_LABELS, type MascotState, type Outfit } from "@/features/mascot";
import { AI_MODELS } from "@/agent/llm";
import { DEFAULT_CONFIG, SCREEN_LABELS, applyConfig, currentConfig, type HubConfig, type Screen } from "@/config";
import { allMistakes, clearMistakes } from "@/lib/mistakes";
import { settings } from "@/lib/store";
import { useTheme } from "@/lib/theme";
import { signOut, tokenStatus } from "@/services/ms";
import { downloadGuides, importGuides } from "@/features/trainer/guides";
import { listPages } from "@/services/confluence";
import { all, useCollection } from "@/lib/store";
import { TemplateGallery } from "@/features/presenter/slides/gallery";

export const Route = createFileRoute("/settings")({
  validateSearch: (s: Record<string, unknown>): { tab?: string } => ({ tab: (s.tab as string) || undefined }),
  component: Settings,
});

const LABELS: Record<string, string> = {
  ...SCREEN_LABELS,
  home: "Home page (a card per screen)",
  aiPanel: "AI panel on every screen",
  visualPreview: "Guides: live preview of the clicked visual",
  visualExplain: 'Guides: "Explain" button',
  visualWhy: 'Guides: "Why did it change?" button (time charts)',
  learnFromMistakes: "Show the AI its past mistakes (kept in this browser)",
  mascot: "Show the mascot in the assistant",
  fileUpload: "Add files to a conversation (paperclip)",
  ratings: "Rate the answers and the conversations (thumbs up / down, optional note)",
  guideEditor: "Guides: edit the guides in the viewer",
  excelUpload: "Excel: upload files into image / file columns",
  guides: "Trainer tab: Guides",
  measures: "Builder tab: Measures",
  visualPdf: "Visual as PDF (Guides)",
  visualPng: "Visual as PNG (Guides)",
  report: "Report export (PDF, .pbix)",
  all: "Every screen",
};

/**
 * Settings (never hidden): AI (this browser's key and model, the admin's instructions), Power BI (the
 * Microsoft account, viewer and export options), General (name, color, theme, screens, options file).
 * Options are applied in this browser with "Apply", and shared through a JSON file on SharePoint.
 */
function Settings() {
  const { tab } = Route.useSearch();
  const navigate = useNavigate({ from: "/settings" });
  const [draft, setDraft] = useState<HubConfig>(currentConfig);
  const [applied, setApplied] = useState(false);
  const set = <K extends keyof HubConfig>(key: K, value: HubConfig[K]) => (setDraft((d) => ({ ...d, [key]: value })), setApplied(false));
  const changed = JSON.stringify(draft) !== JSON.stringify(currentConfig());
  const reports = useCollection("reports");
  return (
    <Workspace
      title="Settings"
      icon={SettingsIcon}
      actions={
        <>
          {applied && !changed && <span className="text-[12.5px] text-success">Applied in this browser.</span>}
          <Button size="sm" variant="ghost" onClick={() => setDraft(structuredClone(DEFAULT_CONFIG))}>
            <RotateCcwIcon /> Defaults
          </Button>
          <Button size="sm" disabled={!changed} onClick={() => (applyConfig(draft), setApplied(true))}>
            <SaveIcon /> Apply
          </Button>
        </>
      }
      tab={tab}
      onTab={(t) => navigate({ search: { tab: t }, replace: true })}
      tabs={[
        {
          id: "ai",
          label: "AI",
          icon: SparklesIcon,
          content: (
            <Page>
              <AiKey />
              <Mistakes learn={draft.learnFromMistakes} onLearn={(v) => set("learnFromMistakes", v)} />
              <Conversations flags={{ ratings: draft.ratings, fileUpload: draft.fileUpload }} onFlags={(v) => (setDraft((d) => ({ ...d, ...v })), setApplied(false))} />
              <Section title="Instructions" description="Added to the AI's own instructions (company vocabulary, tone, rules): for every screen, and per screen.">
                <div className="flex flex-col gap-4">
                  {(Object.keys(draft.prompts) as (keyof HubConfig["prompts"])[]).map((k) => (
                    <div key={k} className="flex flex-col gap-1.5">
                      <Label htmlFor={`prompt-${k}`}>{k === "all" ? "Every screen" : (SCREEN_LABELS[k as Screen] ?? k)}</Label>
                      <Textarea id={`prompt-${k}`} rows={2} value={draft.prompts[k]} onChange={(e) => set("prompts", { ...draft.prompts, [k]: e.target.value })} />
                    </div>
                  ))}
                </div>
              </Section>
            </Page>
          ),
        },
        {
          id: "powerbi",
          label: "Power BI",
          icon: BarChart3Icon,
          content: (
            <Page>
              <Microsoft server={draft.tokenServer} onServer={(v) => set("tokenServer", v)} />
              <Section title="Trainer" description="Tabs and tools of the report viewer (Trainer).">
                <Flags value={{ ...draft.viewerTabs }} onChange={(v) => set("viewerTabs", v)} />
                <div className="mt-3 border-t pt-3">
                  <Flags value={{ visualPreview: draft.visualPreview, visualExplain: draft.visualExplain, visualWhy: draft.visualWhy, guideEditor: draft.guideEditor }} onChange={(v) => (setDraft((d) => ({ ...d, ...v })), setApplied(false))} />
                </div>
              </Section>
              <Section title="Builder" description="Tabs of the Builder.">
                <Flags value={{ measures: draft.builderTabs.measures }} onChange={(v) => set("builderTabs", { ...draft.builderTabs, ...v })} />
              </Section>
              <Section title="Excels">
                <Flags value={{ excelUpload: draft.excelUpload }} onChange={(v) => set("excelUpload", v.excelUpload)} />
              </Section>
              <Section title="Exports" description="The Power BI export API needs a Premium capacity: on Pro, PDF goes through the browser's print dialog.">
                <Flags value={draft.export} onChange={(v) => set("export", v)} />
              </Section>
            </Page>
          ),
        },
        {
          id: "guides",
          label: "Guides",
          icon: BookOpenIcon,
          content: (
            <Page>
              <Guides />
            </Page>
          ),
        },
        {
          id: "slides",
          label: "Slides",
          icon: PresentationIcon,
          content: <TemplateGallery report={reports[0]?.id} />,
        },
        {
          id: "performance",
          label: "Performance",
          icon: GaugeIcon,
          content: (
            <Page>
              <Performance value={draft.performance} onChange={(v) => set("performance", v)} />
            </Page>
          ),
        },
        {
          id: "confluence",
          label: "Confluence",
          icon: FileTextIcon,
          content: (
            <Page>
              <ConfluenceSettings />
            </Page>
          ),
        },
        {
          id: "mascot",
          label: "Mascot",
          icon: SmileIcon,
          content: (
            <Page>
              <MascotPreview />
              <Section title="In the assistant" description="The assistant's mascot: dressed for each screen (Trainer, Builder, Data Scientist, Presenter; the diver elsewhere), thinking or working while the assistant does, talking when it answers, sleeping when nothing happens.">
                <Flags value={{ mascot: draft.mascot }} onChange={(v) => set("mascot", v.mascot)} />
              </Section>
            </Page>
          ),
        },
        {
          id: "general",
          label: "General",
          icon: Settings2Icon,
          content: (
            <Page>
              <Section title="Look">
                <div className="grid gap-4 sm:grid-cols-[minmax(0,1fr)_auto]">
                  <div className="flex flex-col gap-1.5">
                    <Label htmlFor="app-title">App name</Label>
                    <Input id="app-title" value={draft.appTitle} onChange={(e) => set("appTitle", e.target.value)} />
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <Label htmlFor="brand-name">Brand of the slide decks</Label>
                    <Input id="brand-name" value={draft.brandName} onChange={(e) => set("brandName", e.target.value)} />
                    <p className="text-[12px] text-muted-foreground">Default brand name in slide headers and covers (a deck can set its own).</p>
                  </div>
                  <div className="flex flex-col gap-1.5">
                    <Label htmlFor="app-color">Primary color</Label>
                    <span className="flex items-center gap-2">
                      <input id="app-color" type="color" className="h-9 w-11 cursor-pointer rounded-md border" value={/^#[0-9a-f]{6}$/i.test(draft.primaryColor) ? draft.primaryColor : "#2563eb"} onChange={(e) => set("primaryColor", e.target.value)} />
                      <Input className="w-28" value={draft.primaryColor} onChange={(e) => set("primaryColor", e.target.value)} />
                    </span>
                  </div>
                </div>
                <ThemeChoice />
              </Section>
              <Section title="Screens" description="A hidden screen is also closed to direct links. Settings are always shown.">
                <Flags value={draft.screens} onChange={(v) => set("screens", v)} />
                <div className="mt-3 border-t pt-3">
                  <Flags value={{ home: draft.home, aiPanel: draft.aiPanel }} onChange={(v) => (setDraft((d) => ({ ...d, ...v })), setApplied(false))} />
                </div>
              </Section>
            </Page>
          ),
        },
      ]}
    />
  );
}

const Page = ({ children }: { children: ReactNode }) => (
  <div className="min-h-0 flex-1 overflow-y-auto">
    <div className="mx-auto flex max-w-3xl flex-col gap-5 px-6 py-8">{children}</div>
  </div>
);

function Flags<T extends Record<string, boolean>>({ value, onChange }: { value: T; onChange: (v: T) => void }) {
  return (
    <div className="grid gap-x-6 gap-y-3 sm:grid-cols-2">
      {Object.entries(value).map(([k, on]) => (
        <label key={k} className="flex cursor-pointer items-center gap-3 text-[13.5px]">
          <Switch checked={on} onCheckedChange={(v) => onChange({ ...value, [k]: v })} />
          {LABELS[k] ?? k}
        </label>
      ))}
    </div>
  );
}

function ThemeChoice() {
  const { theme, setTheme } = useTheme();
  return (
    <div className="mt-4 flex items-center gap-3 border-t pt-4">
      <span className="text-[13.5px] font-medium">Theme (this browser)</span>
      <ToggleGroup type="single" variant="outline" value={theme} onValueChange={(v) => v && setTheme(v as typeof theme)}>
        <ToggleGroupItem value="light" aria-label="Light" className="gap-1.5 px-3"><SunIcon /> Light</ToggleGroupItem>
        <ToggleGroupItem value="dark" aria-label="Dark" className="gap-1.5 px-3"><MoonIcon /> Dark</ToggleGroupItem>
        <ToggleGroupItem value="system" aria-label="System" className="gap-1.5 px-3"><MonitorIcon /> System</ToggleGroupItem>
      </ToggleGroup>
    </div>
  );
}

/* -------------------------------- Microsoft ------------------------------- */

type Status = Awaited<ReturnType<typeof tokenStatus>>;

/**
 * The Microsoft account: the sign-in lives in the local token server (bun dev starts it); the app only
 * asks it for tokens. While someone has to sign in, its URL and code are shown (read again every few seconds).
 */
function Microsoft({ server, onServer }: { server: string; onServer: (v: string) => void }) {
  const [status, setStatus] = useState<Status>();
  useEffect(() => {
    let stop = false;
    const poll = async () => {
      const s = await tokenStatus();
      if (stop) return;
      setStatus(s);
      if (!s.signedIn) setTimeout(poll, 4000);
    };
    void poll();
    return () => void (stop = true);
  }, []);
  return (
    <Section title="Microsoft account" description="One sign-in (device code, any organization), kept by the local token server: it gives the app the tokens of Power BI, Fabric and Graph.">
      {!status ? (
        <p className="text-muted-foreground">Reading…</p>
      ) : status.signedIn ? (
        <div className="flex items-center gap-3">
          <span className="flex size-9 items-center justify-center rounded-full bg-primary/10 font-semibold text-primary">{status.account?.slice(0, 1).toUpperCase()}</span>
          <span className="flex-1 text-[13.5px]">
            Signed in as <b>{status.account}</b>
          </span>
          <Button variant="outline" size="sm" onClick={() => void signOut().then(() => setStatus(undefined)).then(() => tokenStatus().then(setStatus))}>
            <LogOutIcon /> Sign out
          </Button>
        </div>
      ) : status.signIn ? (
        <p className="text-[13.5px]">
          Open{" "}
          <a className="text-primary underline" href={status.signIn.url} target="_blank" rel="noreferrer">
            {status.signIn.url}
          </a>{" "}
          and enter the code <b className="font-mono text-[17px]">{status.signIn.code}</b>. This page follows the sign-in by itself.
        </p>
      ) : (
        <p className="text-[13.5px] text-destructive">{status.error}</p>
      )}
      <div className="mt-4 flex flex-col gap-1.5 border-t pt-4">
        <Label htmlFor="token-server">Token server</Label>
        <Input id="token-server" className="font-mono text-[13px]" value={server} onChange={(e) => onServer(e.target.value)} />
        <p className="text-[12px] text-muted-foreground">Started with the app by "bun dev" (or alone: "bun run token", or the dist/hub-token-server program).</p>
      </div>
    </Section>
  );
}

/* --------------------------------- Guides --------------------------------- */

/** The guides of the report viewer, kept in this browser: download them as Excel, import Excel or CSV. */
function Guides() {
  const guides = useCollection("guides");
  const file = useRef<HTMLInputElement>(null);
  const [state, setState] = useState<{ busy?: string; done?: string; error?: string }>({});
  const run = (busy: string, task: () => Promise<string>) => (setState({ busy }), task().then((done) => setState({ done }), (e) => setState({ error: e instanceof Error ? e.message : String(e) })));
  return (
    <Section
      title="Guides"
      description={`The guide of each visual of the reports (Trainer), kept in this browser: ${guides.length} guide(s). The Excel / CSV file has the columns id, report_id, page, visual, title, content.`}
      actions={
        <>
          <input ref={file} type="file" accept=".xlsx,.csv" hidden onChange={(e) => {
            const f = e.target.files?.[0];
            e.target.value = "";
            if (f) run("Importing…", async () => `${await importGuides(f)} guide(s) imported.`);
          }} />
          <Button size="sm" variant="outline" onClick={() => file.current?.click()}>
            <UploadIcon /> Import Excel / CSV
          </Button>
          <Button size="sm" variant="outline" disabled={!guides.length} onClick={() => run("Writing the file…", async () => (await downloadGuides(), "Downloaded."))}>
            <DownloadIcon /> Download as Excel
          </Button>
        </>
      }
    >
      {state.busy && <p className="text-[13px] text-muted-foreground">{state.busy}</p>}
      {state.done && <p className="text-[13px] text-success">{state.done}</p>}
      {state.error && <p className="text-[13px] text-destructive">{state.error}</p>}
    </Section>
  );
}

/* ------------------------------- Performance ------------------------------ */

const PERF_FIELDS: { key: keyof HubConfig["performance"]; label: string; hint: string }[] = [
  { key: "maxRows", label: "Rows per query (max)", hint: "Rows a DAX query may return: more are cut (the AI knows it)." },
  { key: "timeoutMs", label: "Query time limit (ms)", hint: "A query running longer is stopped." },
  { key: "concurrency", label: "Queries at once per model", hint: "The others wait their turn." },
  { key: "intervalMs", label: "Pause between queries (ms)", hint: "On the same model." },
  { key: "cacheMs", label: "Same query reused for (ms)", hint: "An identical query's result is reused this long." },
  { key: "maxEmbeds", label: "Reports embedded at once", hint: "Power BI pages loaded together (slides)." },
];

/** The limits that keep the semantic models fast for everyone (docs/performance.md). */
function Performance({ value, onChange }: { value: HubConfig["performance"]; onChange: (v: HubConfig["performance"]) => void }) {
  return (
    <Section title="Limits" description="They protect the semantic models when many people use the app on the same model. The AI is told the row limit.">
      <div className="grid gap-4 sm:grid-cols-2">
        {PERF_FIELDS.map((f) => (
          <div key={f.key} className="flex flex-col gap-1.5">
            <Label htmlFor={`perf-${f.key}`}>{f.label}</Label>
            <Input id={`perf-${f.key}`} type="number" min={0} value={value[f.key]} onChange={(e) => onChange({ ...value, [f.key]: Math.max(0, Number(e.target.value) || 0) })} />
            <p className="text-[12px] text-muted-foreground">{f.hint}</p>
          </div>
        ))}
      </div>
    </Section>
  );
}

/* -------------------------------- Confluence -------------------------------- */

/**
 * Where the Confluence screen reads and writes its pages: the address of the API route, a key, and the
 * folder (a page or folder id) whose pages it lists. Kept in this browser only (the key is a secret).
 */
function ConfluenceSettings() {
  const [values, setValues] = useState({ url: "", key: "", folderId: "" });
  const [state, setState] = useState<{ busy?: boolean; done?: string; error?: string }>({});
  useEffect(() => void settings.get().then((s) => s.confluence && setValues(s.confluence)), []);
  const set = (v: Partial<typeof values>) => (setValues({ ...values, ...v }), setState({}));
  const save = async (test: boolean) => {
    setState({ busy: true });
    try {
      await settings.set({ confluence: { url: values.url.trim(), key: values.key.trim(), folderId: values.folderId.trim() } });
      setState({ done: test ? `Connected: ${(await listPages()).length} page(s) in the folder.` : "Saved." });
    } catch (e) {
      setState({ error: e instanceof Error ? e.message : String(e) });
    }
  };
  return (
    <Section
      title="Confluence"
      description="Kept in this browser only. The Confluence screen (Settings › General › screens) lists the pages of this folder and edits them as markdown."
      actions={
        <>
          {state.done && <span className="text-[12.5px] text-success">{state.done}</span>}
          {state.error && <span className="max-w-80 text-[12.5px] text-destructive">{state.error}</span>}
          <Button size="sm" variant="outline" disabled={state.busy} onClick={() => void save(true)}>
            Test
          </Button>
          <Button size="sm" disabled={state.busy} onClick={() => void save(false)}>
            <SaveIcon /> Save
          </Button>
        </>
      }
    >
      <div className="grid gap-4 sm:grid-cols-2">
        <div className="flex flex-col gap-1.5 sm:col-span-2">
          <Label htmlFor="confluence-url">API route (URL)</Label>
          <Input id="confluence-url" placeholder="https://company.atlassian.net/wiki" value={values.url} onChange={(e) => set({ url: e.target.value })} />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="confluence-key">API key</Label>
          <Input id="confluence-key" type="password" placeholder="email:token (Cloud), or a personal access token" value={values.key} onChange={(e) => set({ key: e.target.value })} />
        </div>
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="confluence-folder">Folder id</Label>
          <Input id="confluence-folder" placeholder="123456789" value={values.folderId} onChange={(e) => set({ folderId: e.target.value })} />
        </div>
      </div>
    </Section>
  );
}

/* ----------------------------------- AI ----------------------------------- */

function AiKey() {
  const [values, setValues] = useState({ aiKey: "", model: AI_MODELS[0] as string, implementerModel: AI_MODELS[0] as string, maxSteps: 24 });
  const [saved, setSaved] = useState(false);
  useEffect(() => void settings.get().then((s) => setValues({ aiKey: s.aiKey ?? "", model: s.model ?? AI_MODELS[0], implementerModel: s.implementerModel ?? s.model ?? AI_MODELS[0], maxSteps: s.maxSteps ?? 24 })), []);
  const set = (v: Partial<typeof values>) => (setValues({ ...values, ...v }), setSaved(false));
  const select = (id: string, label: string, key: "model" | "implementerModel") => (
    <div className="flex flex-col gap-1.5">
      <Label htmlFor={id}>{label}</Label>
      <select id={id} className="input h-9" value={values[key]} onChange={(e) => set({ [key]: e.target.value })}>
        {AI_MODELS.map((m) => (
          <option key={m}>{m}</option>
        ))}
      </select>
    </div>
  );
  return (
    <Section
      title="Key and models"
      description="Kept in this browser only: the AI API is called from the browser with this key. The planner explores and asks (few calls); the implementer makes the changes (many calls)."
      actions={
        <>
          {saved && <span className="text-[12.5px] text-success">Saved.</span>}
          <Button size="sm" onClick={() => void settings.set(values).then(() => setSaved(true))}>
            <SaveIcon /> Save
          </Button>
        </>
      }
    >
      <div className="grid gap-4 sm:grid-cols-[minmax(0,1fr)_200px]">
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="ai-key">AI key (token, or NAME=value lines)</Label>
          <Textarea id="ai-key" rows={3} className="font-mono text-[13px]" value={values.aiKey} onChange={(e) => set({ aiKey: e.target.value })} />
        </div>
        <div className="flex flex-col gap-3">
          {select("ai-model", "Planner model", "model")}
          {select("ai-implementer-model", "Implementer model", "implementerModel")}
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="ai-max-steps">Maximum steps</Label>
            <Input id="ai-max-steps" type="number" min={4} max={100} value={values.maxSteps} onChange={(e) => set({ maxSteps: Math.max(4, Math.min(100, Number(e.target.value) || 24)) })} />
            <span className="text-[11.5px] text-muted-foreground">Tool turns of the implementer per request (the planner gets half).</span>
          </div>
        </div>
      </div>
    </Section>
  );
}

/** The AI's failed tool calls kept in this browser: shown to it as mistakes not to repeat; downloaded to improve prompts and tools. */
function Mistakes({ learn, onLearn }: { learn: boolean; onLearn: (v: boolean) => void }) {
  const [count, setCount] = useState<number>();
  useEffect(() => void allMistakes().then((m) => setCount(m.length)), []);
  const download = async () => {
    const json = JSON.stringify({ exportedAt: new Date().toISOString(), mistakes: await allMistakes() }, null, 2);
    const url = URL.createObjectURL(new Blob([json], { type: "application/json" }));
    Object.assign(document.createElement("a"), { href: url, download: `ai-mistakes-${new Date().toISOString().slice(0, 10)}.json` }).click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  };
  return (
    <Section
      title="Mistakes"
      description="Each failed tool call of the AI, with the call that fixed it. The latest ones of a screen are shown to the AI as examples not to follow."
      actions={
        <>
          <Button size="sm" variant="outline" disabled={!count} onClick={() => void download()}>
            <DownloadIcon /> Download ({count ?? "…"})
          </Button>
          <Button size="sm" variant="ghost" disabled={!count} onClick={() => void clearMistakes().then(() => setCount(0))}>
            Clear
          </Button>
        </>
      }
    >
      <Flags value={{ learnFromMistakes: learn }} onChange={(v) => onLearn(v.learnFromMistakes)} />
    </Section>
  );
}

/** Every conversation of this browser, searched in a table (screen, item, ratings); the ratings option. */
const PERIODS = { all: "Since the beginning", today: "Today", week: "This week", month: "This month" } as const;

/** Every conversation of this browser: searched in a table, or downloaded (all, or since a date); the ratings and file options. */
function Conversations({ flags, onFlags }: { flags: { ratings: boolean; fileUpload: boolean }; onFlags: (v: { ratings: boolean; fileUpload: boolean }) => void }) {
  const [open, setOpen] = useState(false);
  const [period, setPeriod] = useState<keyof typeof PERIODS>("all");
  const download = async () => {
    const now = new Date();
    const start = { all: 0, today: new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime(), week: now.getTime() - 7 * 86_400_000, month: new Date(now.getFullYear(), now.getMonth(), 1).getTime() }[period];
    const list = (await all("conversations")).filter((c) => c.updatedAt >= start);
    const url = URL.createObjectURL(new Blob([JSON.stringify({ exportedAt: now.toISOString(), period, conversations: list }, null, 2)], { type: "application/json" }));
    Object.assign(document.createElement("a"), { href: url, download: `conversations-${period}-${now.toISOString().slice(0, 10)}.json` }).click();
    setTimeout(() => URL.revokeObjectURL(url), 10_000);
  };
  return (
    <Section
      title="Conversations"
      description="Search every conversation with the assistant (its screen, the item it is about, its ratings and notes), or download them."
      actions={
        <>
          <select className="input h-8 w-40" value={period} onChange={(e) => setPeriod(e.target.value as keyof typeof PERIODS)} aria-label="Conversations since">
            {(Object.keys(PERIODS) as (keyof typeof PERIODS)[]).map((p) => (
              <option key={p} value={p}>
                {PERIODS[p]}
              </option>
            ))}
          </select>
          <Button size="sm" variant="outline" onClick={() => void download()}>
            <DownloadIcon /> Download
          </Button>
          <Button size="sm" variant="outline" onClick={() => setOpen(true)}>
            <SearchIcon /> Search
          </Button>
        </>
      }
    >
      <Flags value={flags} onChange={onFlags} />
      <ConversationsDialog open={open} onOpenChange={setOpen} />
    </Section>
  );
}

/** The mascot in each outfit and animation, to see them. */
function MascotPreview() {
  const [outfit, setOutfit] = useState<Outfit>("diver");
  const [state, setState] = useState<MascotState>("idle");
  return (
    <Section title="Mascot" description="Choose an outfit and an animation to see them.">
      <div className="grid gap-6 sm:grid-cols-[200px_minmax(0,1fr)]">
        <div className="flex flex-col gap-4">
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="mascot-outfit">Outfit</Label>
            <select id="mascot-outfit" className="input h-9" value={outfit} onChange={(e) => setOutfit(e.target.value as Outfit)}>
              {(Object.keys(OUTFIT_LABELS) as Outfit[]).map((o) => (
                <option key={o} value={o}>
                  {OUTFIT_LABELS[o]}
                </option>
              ))}
            </select>
          </div>
          <div className="flex flex-col gap-1.5">
            <Label htmlFor="mascot-state">Animation</Label>
            <select id="mascot-state" className="input h-9" value={state} onChange={(e) => setState(e.target.value as MascotState)}>
              {STATES.map((s) => (
                <option key={s} value={s}>
                  {STATE_LABELS[s]}
                </option>
              ))}
            </select>
          </div>
        </div>
        <div className="flex items-center justify-center rounded-xl border bg-muted/30 p-6">
          <Mascot outfit={outfit} state={state} size={240} />
        </div>
      </div>
    </Section>
  );
}

