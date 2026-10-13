import { useEffect, useState, type ReactNode } from "react";
import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { DownloadIcon, LogOutIcon, RotateCcwIcon, SaveIcon, UploadIcon } from "lucide-react";
import { AI_MODELS } from "@/agent/llm";
import { DEFAULT_CONFIG, applyConfig, currentConfig, type HubConfig } from "@/config";
import { loadHubConfig, saveHubConfig } from "@/lib/hub-config";
import { errorText } from "@/lib/ms";
import { settings } from "@/lib/store";
import { authStatus, signIn, signOut } from "@/server/ms";

type Tab = "config" | "ai" | "microsoft";

export const Route = createFileRoute("/settings")({
  validateSearch: (s: Record<string, unknown>): { tab?: Tab } => ({ tab: ["config", "ai", "microsoft"].includes(s.tab as string) ? (s.tab as Tab) : undefined }),
  component: Settings,
});

const TABS: { id: Tab; label: string }[] = [
  { id: "config", label: "Config" },
  { id: "ai", label: "AI" },
  { id: "microsoft", label: "Microsoft" },
];

/** Settings (never hidden): the app's options for everyone (SharePoint file), this browser's AI key, the Microsoft account. */
function Settings() {
  const { tab = "config" } = Route.useSearch();
  const navigate = useNavigate({ from: "/settings" });
  return (
    <div className="flex h-full flex-col">
      <div className="flex h-12 shrink-0 items-center gap-1 border-b bg-card px-3">
        {TABS.map((t) => (
          <button key={t.id} type="button" className="tab" data-active={t.id === tab} onClick={() => navigate({ search: { tab: t.id }, replace: true })}>
            {t.label}
          </button>
        ))}
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto">
        <div className="mx-auto flex max-w-3xl flex-col gap-4 p-6">{tab === "config" ? <ConfigTab /> : tab === "ai" ? <AiTab /> : <MicrosoftTab />}</div>
      </div>
    </div>
  );
}

const Card = ({ title, hint, children }: { title: string; hint?: string; children: ReactNode }) => (
  <section className="flex flex-col gap-2 rounded-lg border bg-card p-4">
    <h2 className="font-medium">{title}</h2>
    {hint && <p className="text-[12px] text-muted-foreground">{hint}</p>}
    {children}
  </section>
);

/* --------------------------------- Config --------------------------------- */

/** Labels of the options (the keys of config.ts). */
const LABELS: Record<string, string> = {
  showHeader: "Top bar",
  aiPanel: "AI panel (every screen)",
  viewerEditor: "Viewer: guides and links can be edited",
  visualPreview: "Viewer › Guides: preview of the clicked visual",
  visualExplain: 'Viewer › Guides: "Explain" button',
  library: "Library",
  viewer: "Report viewer",
  builder: "Report builder",
  datasets: "Dataset builder",
  dax: "DAX builder",
  pivot: "TCD builder",
  info: "Info",
  guide: "Guides",
  ai: "AI",
  inspiration: "Inspiration",
  schema: "Schema",
  fields: "Fields",
  visualPdf: "Visual as PDF (Guides)",
  visualPng: "Visual as PNG (Guides)",
  report: "Report export (PDF, .pbix)",
  all: "Every screen",
};
const label = (key: string) => LABELS[key] ?? key;

function Flags<T extends Record<string, boolean>>({ value, onChange }: { value: T; onChange: (v: T) => void }) {
  return (
    <div className="grid grid-cols-2 gap-x-4 gap-y-1.5">
      {Object.entries(value).map(([k, on]) => (
        <label key={k} className="flex items-center gap-2 text-[13px]">
          <input type="checkbox" className="accent-foreground" checked={on} onChange={(e) => onChange({ ...value, [k]: e.target.checked })} />
          {label(k)}
        </label>
      ))}
    </div>
  );
}

function ConfigTab() {
  const [draft, setDraft] = useState<HubConfig>(currentConfig);
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: string; done?: string; error?: string }>({});
  useEffect(() => void settings.get().then((s) => setLink(s.configLink ?? "")), []);
  const set = <K extends keyof HubConfig>(key: K, value: HubConfig[K]) => setDraft((d) => ({ ...d, [key]: value }));
  const run = (busy: string, task: () => Promise<string>) => {
    setState({ busy });
    task().then((done) => setState({ done }), (e) => setState({ error: errorText(e) }));
  };
  const general = { showHeader: draft.showHeader, aiPanel: draft.aiPanel, viewerEditor: draft.viewerEditor, visualPreview: draft.visualPreview, visualExplain: draft.visualExplain };
  return (
    <>
      <Card title="Configuration file (SharePoint)" hint="The options below are shared through a JSON file on SharePoint / OneDrive: everyone whose browser has this link gets them at start. Give a file, or a folder (hub-config.json is created in it).">
        <input className="input" placeholder="Sharing link of hub-config.json, or of a folder" value={link} onChange={(e) => setLink(e.target.value)} />
        <div className="flex flex-wrap gap-1.5">
          <button
            type="button"
            className="btn"
            disabled={!link.trim() || !!state.busy}
            onClick={() =>
              run("Reading…", async () => {
                await loadHubConfig(link.trim());
                await settings.set({ configLink: link.trim() });
                return "Options read from the file and applied.";
              })
            }
          >
            <DownloadIcon /> Load from the file
          </button>
          <button
            type="button"
            className="btn-primary"
            disabled={!link.trim() || !!state.busy}
            onClick={() =>
              run("Saving…", async () => {
                applyConfig(draft);
                const saved = await saveHubConfig(link.trim());
                setLink(saved);
                return "Saved to SharePoint and applied.";
              })
            }
          >
            <UploadIcon /> Save to SharePoint
          </button>
        </div>
        {state.busy && <p className="text-[12px] text-muted-foreground">{state.busy}</p>}
        {state.done && <p className="text-[12px] text-green-700">{state.done}</p>}
        {state.error && <p className="text-[12px] break-words text-destructive">{state.error}</p>}
      </Card>

      <Card title="General">
        <div className="grid grid-cols-[1fr_auto] gap-2">
          <label className="flex flex-col gap-1">
            <span className="label">App name</span>
            <input className="input" value={draft.appTitle} onChange={(e) => set("appTitle", e.target.value)} />
          </label>
          <label className="flex flex-col gap-1">
            <span className="label">Primary color</span>
            <span className="flex items-center gap-1.5">
              <input type="color" className="h-8 w-10 rounded border" value={/^#[0-9a-f]{6}$/i.test(draft.primaryColor) ? draft.primaryColor : "#18181b"} onChange={(e) => set("primaryColor", e.target.value)} />
              <input className="input w-28" value={draft.primaryColor} onChange={(e) => set("primaryColor", e.target.value)} />
            </span>
          </label>
        </div>
        <Flags value={general} onChange={(v) => setDraft((d) => ({ ...d, ...v }))} />
      </Card>

      <Card title="Screens" hint="A hidden screen leaves the menu and its links open the first screen shown.">
        <Flags value={draft.screens} onChange={(v) => set("screens", v)} />
      </Card>

      <Card title="Panel tabs">
        {([["viewerTabs", "Report viewer"], ["builderTabs", "Report builder"], ["datasetTabs", "Dataset builder"], ["daxTabs", "DAX builder"], ["pivotTabs", "TCD builder"]] as const).map(([key, title]) => (
          <div key={key} className="flex flex-col gap-1">
            <span className="label">{title}</span>
            <Flags value={draft[key] as Record<string, boolean>} onChange={(v) => set(key, v as never)} />
          </div>
        ))}
      </Card>

      <Card title="Exports">
        <Flags value={draft.export} onChange={(v) => set("export", v)} />
      </Card>

      <Card title="AI instructions" hint="Added to the AI's own instructions: company vocabulary, tone, rules (e.g. &quot;CA = [Revenue]&quot;, &quot;answer in French&quot;).">
        {(Object.keys(draft.prompts) as (keyof HubConfig["prompts"])[]).map((k) => (
          <label key={k} className="flex flex-col gap-1">
            <span className="label">{label(k)}</span>
            <textarea className="input h-16 py-1" value={draft.prompts[k]} onChange={(e) => set("prompts", { ...draft.prompts, [k]: e.target.value })} />
          </label>
        ))}
      </Card>

      <div className="sticky bottom-0 flex justify-end gap-1.5 border-t bg-background/95 py-3 backdrop-blur">
        <button type="button" className="btn" onClick={() => setDraft(structuredClone(DEFAULT_CONFIG))}>
          <RotateCcwIcon /> Defaults
        </button>
        <button type="button" className="btn-primary" onClick={() => (applyConfig(draft), setState({ done: "Applied in this browser (save to SharePoint to share it)." }))}>
          <SaveIcon /> Apply here
        </button>
      </div>
    </>
  );
}

/* ----------------------------------- AI ----------------------------------- */

function AiTab() {
  const [values, setValues] = useState({ aiKey: "", model: AI_MODELS[0] as string });
  const [saved, setSaved] = useState(false);
  useEffect(() => void settings.get().then((s) => setValues({ aiKey: s.aiKey ?? "", model: s.model ?? AI_MODELS[0] })), []);
  return (
    <Card title="AI" hint="Kept in this browser only. The AI API is called from the browser with this key.">
      <label className="label">AI key (token, or NAME=value lines)</label>
      <textarea className="input h-20 py-1 font-mono" value={values.aiKey} onChange={(e) => (setValues({ ...values, aiKey: e.target.value }), setSaved(false))} />
      <label className="label">Model</label>
      <select className="input" value={values.model} onChange={(e) => (setValues({ ...values, model: e.target.value }), setSaved(false))}>
        {AI_MODELS.map((m) => (
          <option key={m}>{m}</option>
        ))}
      </select>
      <div className="flex items-center justify-end gap-2">
        {saved && <span className="text-[12px] text-green-700">Saved.</span>}
        <button type="button" className="btn-primary" onClick={() => void settings.set(values).then(() => setSaved(true))}>
          <SaveIcon /> Save
        </button>
      </div>
    </Card>
  );
}

/* -------------------------------- Microsoft ------------------------------- */

type Status = Awaited<ReturnType<typeof authStatus>>;

function MicrosoftTab() {
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
    <Card title="Microsoft account" hint="One sign-in (device code, any organization) gives access to Power BI, Fabric, SharePoint / OneDrive and Forms. The session stays on the server of this app.">
      {!status ? (
        <p className="text-[12px] text-muted-foreground">Reading…</p>
      ) : status.signedIn ? (
        <div className="flex items-center gap-2">
          <span className="flex-1">
            Signed in as <b>{status.account}</b>
          </span>
          <button type="button" className="btn" onClick={() => void signOut().then(() => setStatus({ signedIn: false, account: undefined, pending: null }))}>
            <LogOutIcon /> Sign out
          </button>
        </div>
      ) : status.pending ? (
        <p className="text-[13px]">
          Open{" "}
          <a className="underline" href={status.pending.verificationUri} target="_blank" rel="noreferrer">
            {status.pending.verificationUri}
          </a>{" "}
          and enter the code <b className="font-mono text-[15px]">{status.pending.userCode}</b>
          {status.pending.error && <span className="text-destructive"> — {status.pending.error}</span>}
        </p>
      ) : (
        <button type="button" className="btn-primary w-fit" onClick={() => void signIn().then((pending) => setStatus({ signedIn: false, pending, account: undefined }))}>
          Sign in with Microsoft
        </button>
      )}
    </Card>
  );
}
