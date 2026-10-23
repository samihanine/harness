import { useEffect, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { BarChart3Icon, DownloadIcon, LogOutIcon, MonitorIcon, MoonIcon, RotateCcwIcon, SaveIcon, Settings2Icon, SettingsIcon, SparklesIcon, SunIcon, UploadIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Textarea } from "@/components/ui/textarea";
import { ToggleGroup, ToggleGroupItem } from "@/components/ui/toggle-group";
import { Section, StatusText, Workspace } from "@/components/workspace";
import { AI_MODELS } from "@/agent/llm";
import { DEFAULT_CONFIG, SCREEN_LABELS, applyConfig, currentConfig, type HubConfig } from "@/config";
import { loadHubConfig, saveHubConfig } from "@/lib/hub-config";
import { errorText } from "@/lib/ms";
import { settings } from "@/lib/store";
import { useTheme } from "@/lib/theme";
import { authStatus, signIn, signOut } from "@/server/ms";

export const Route = createFileRoute("/settings")({
  validateSearch: (s: Record<string, unknown>): { tab?: string } => ({ tab: (s.tab as string) || undefined }),
  component: Settings,
});

const LABELS: Record<string, string> = {
  ...SCREEN_LABELS,
  aiPanel: "AI panel on every screen",
  visualPreview: "Guides: live preview of the clicked visual",
  visualExplain: 'Guides: "Explain" button',
  guideEditor: "Guides: edit the guides in the viewer",
  excelUpload: "Excel: upload files into image / file columns",
  guides: "Viewer tab: Guides",
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
              <Section title="Instructions" description="Added to the AI's own instructions (company vocabulary, tone, rules): for every screen, and per screen.">
                <div className="flex flex-col gap-4">
                  {(Object.keys(draft.prompts) as (keyof HubConfig["prompts"])[]).map((k) => (
                    <div key={k} className="flex flex-col gap-1.5">
                      <Label htmlFor={`prompt-${k}`}>{LABELS[k] ?? k}</Label>
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
              <Microsoft />
              <Section title="Viewer" description="Tabs and tools of the report viewer.">
                <Flags value={{ ...draft.viewerTabs }} onChange={(v) => set("viewerTabs", v)} />
                <div className="mt-3 border-t pt-3">
                  <Flags value={{ visualPreview: draft.visualPreview, visualExplain: draft.visualExplain, guideEditor: draft.guideEditor }} onChange={(v) => (setDraft((d) => ({ ...d, ...v })), setApplied(false))} />
                </div>
              </Section>
              <Section title="Exports" description="The Power BI export API needs a Premium capacity: on Pro, PDF goes through the browser's print dialog.">
                <Flags value={draft.export} onChange={(v) => set("export", v)} />
              </Section>
              <Section title="Excel">
                <Flags value={{ excelUpload: draft.excelUpload }} onChange={(v) => set("excelUpload", v.excelUpload)} />
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
                  <Flags value={{ aiPanel: draft.aiPanel }} onChange={(v) => set("aiPanel", v.aiPanel)} />
                </div>
              </Section>
              <ConfigFile draft={draft} onLoaded={() => setDraft(currentConfig())} />
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

type Status = Awaited<ReturnType<typeof authStatus>>;

function Microsoft() {
  const [status, setStatus] = useState<Status>();
  useEffect(() => {
    let stop = false;
    const poll = async () => {
      const s = await authStatus().catch(() => undefined);
      if (stop) return;
      setStatus(s);
      if (s?.pending) setTimeout(poll, 3000);
    };
    void poll();
    return () => void (stop = true);
  }, [status?.pending?.userCode]);
  return (
    <Section title="Microsoft account" description="One sign-in (device code, any organization) gives access to Power BI, Fabric, SharePoint / OneDrive and Forms.">
      {!status ? (
        <p className="text-muted-foreground">Reading…</p>
      ) : status.signedIn ? (
        <div className="flex items-center gap-3">
          <span className="flex size-9 items-center justify-center rounded-full bg-primary/10 font-semibold text-primary">{status.account?.slice(0, 1).toUpperCase()}</span>
          <span className="flex-1 text-[13.5px]">
            Signed in as <b>{status.account}</b>
          </span>
          <Button variant="outline" size="sm" onClick={() => void signOut().then(() => setStatus({ signedIn: false, account: undefined, pending: null }))}>
            <LogOutIcon /> Sign out
          </Button>
        </div>
      ) : status.pending ? (
        <p className="text-[13.5px]">
          Open{" "}
          <a className="text-primary underline" href={status.pending.verificationUri} target="_blank" rel="noreferrer">
            {status.pending.verificationUri}
          </a>{" "}
          and enter the code <b className="font-mono text-[17px]">{status.pending.userCode}</b>
          {status.pending.error && <span className="text-destructive"> — {status.pending.error}</span>}
        </p>
      ) : (
        <Button onClick={() => void signIn().then((pending) => setStatus({ signedIn: false, pending, account: undefined }))}>Sign in with Microsoft</Button>
      )}
    </Section>
  );
}

/* ----------------------------------- AI ----------------------------------- */

function AiKey() {
  const [values, setValues] = useState({ aiKey: "", model: AI_MODELS[0] as string });
  const [saved, setSaved] = useState(false);
  useEffect(() => void settings.get().then((s) => setValues({ aiKey: s.aiKey ?? "", model: s.model ?? AI_MODELS[0] })), []);
  const set = (v: Partial<typeof values>) => (setValues({ ...values, ...v }), setSaved(false));
  return (
    <Section
      title="Key and model"
      description="Kept in this browser only: the AI API is called from the browser with this key."
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
        <div className="flex flex-col gap-1.5">
          <Label htmlFor="ai-model">Model</Label>
          <select id="ai-model" className="input h-9" value={values.model} onChange={(e) => set({ model: e.target.value })}>
            {AI_MODELS.map((m) => (
              <option key={m}>{m}</option>
            ))}
          </select>
        </div>
      </div>
    </Section>
  );
}

/* --------------------------------- Config file --------------------------------- */

/** The options for everyone: a JSON file on SharePoint, read at start by every browser that has its link. */
function ConfigFile({ draft, onLoaded }: { draft: HubConfig; onLoaded: () => void }) {
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: string; done?: string; error?: string }>({});
  useEffect(() => void settings.get().then((s) => setLink(s.configLink ?? "")), []);
  const run = (busy: string, task: () => Promise<string>) => {
    setState({ busy });
    task().then((done) => setState({ done }), (e) => setState({ error: errorText(e) }));
  };
  return (
    <Section title="Options file" description="Share these options: save them to a JSON file on SharePoint / OneDrive; every browser with its link reads it at start." actions={<StatusText state={state} />}>
      <div className="flex gap-2">
        <Input placeholder="Link of hub-config.json, or of a folder (the file is created in it)" value={link} onChange={(e) => setLink(e.target.value)} />
        <Button
          variant="outline"
          disabled={!link.trim() || !!state.busy}
          onClick={() =>
            run("Reading…", async () => {
              await loadHubConfig(link.trim());
              await settings.set({ configLink: link.trim() });
              onLoaded();
              return "Options read from the file and applied.";
            })
          }
        >
          <DownloadIcon /> Load
        </Button>
        <Button
          variant="outline"
          disabled={!link.trim() || !!state.busy}
          onClick={() =>
            run("Saving…", async () => {
              applyConfig(draft);
              setLink(await saveHubConfig(link.trim()));
              return "Saved to SharePoint and applied.";
            })
          }
        >
          <UploadIcon /> Save to the file
        </Button>
      </div>
    </Section>
  );
}
