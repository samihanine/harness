import { useEffect, useState, type ReactNode } from "react";
import { createFileRoute } from "@tanstack/react-router";
import { DownloadIcon, LogOutIcon, RotateCcwIcon, SaveIcon, UploadIcon } from "lucide-react";
import { AI_MODELS } from "@/agent/llm";
import { DEFAULT_CONFIG, SCREEN_LABELS, applyConfig, currentConfig, type HubConfig } from "@/config";
import { loadHubConfig, saveHubConfig } from "@/lib/hub-config";
import { errorText } from "@/lib/ms";
import { settings } from "@/lib/store";
import { authStatus, signIn, signOut } from "@/server/ms";

export const Route = createFileRoute("/settings")({ component: Settings });

/** Settings (never hidden), one page: the Microsoft account, this browser's AI key, the app's options for everyone. */
function Settings() {
  return (
    <div className="h-full overflow-y-auto">
      <div className="mx-auto flex max-w-3xl flex-col gap-5 px-6 py-8">
        <h1 className="text-[26px] font-semibold">Settings</h1>
        <MicrosoftCard />
        <AiCard />
        <ConfigCard />
      </div>
    </div>
  );
}

const Card = ({ title, hint, children }: { title: string; hint?: string; children: ReactNode }) => (
  <section className="card flex flex-col gap-3 p-5">
    <div>
      <h2 className="text-[16px] font-semibold">{title}</h2>
      {hint && <p className="mt-0.5 text-[13px] text-muted-foreground">{hint}</p>}
    </div>
    {children}
  </section>
);

/* -------------------------------- Microsoft ------------------------------- */

type Status = Awaited<ReturnType<typeof authStatus>>;

function MicrosoftCard() {
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
    <Card title="Microsoft account" hint="One sign-in (device code, any organization) gives access to Power BI, Fabric, SharePoint / OneDrive and Forms.">
      {!status ? (
        <p className="text-muted-foreground">Reading…</p>
      ) : status.signedIn ? (
        <div className="flex items-center gap-3">
          <span className="flex size-9 items-center justify-center rounded-full bg-primary/10 font-semibold text-primary">{status.account?.slice(0, 1).toUpperCase()}</span>
          <span className="flex-1">
            Signed in as <b>{status.account}</b>
          </span>
          <button type="button" className="btn" onClick={() => void signOut().then(() => setStatus({ signedIn: false, account: undefined, pending: null }))}>
            <LogOutIcon /> Sign out
          </button>
        </div>
      ) : status.pending ? (
        <p>
          Open{" "}
          <a className="text-primary underline" href={status.pending.verificationUri} target="_blank" rel="noreferrer">
            {status.pending.verificationUri}
          </a>{" "}
          and enter the code <b className="font-mono text-[17px]">{status.pending.userCode}</b>
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

/* ----------------------------------- AI ----------------------------------- */

function AiCard() {
  const [values, setValues] = useState({ aiKey: "", model: AI_MODELS[0] as string });
  const [saved, setSaved] = useState(false);
  useEffect(() => void settings.get().then((s) => setValues({ aiKey: s.aiKey ?? "", model: s.model ?? AI_MODELS[0] })), []);
  const set = (v: Partial<typeof values>) => (setValues({ ...values, ...v }), setSaved(false));
  return (
    <Card title="AI" hint="Kept in this browser only: the AI API is called from the browser with this key.">
      <div className="grid grid-cols-[1fr_200px] gap-3">
        <label className="flex flex-col gap-1.5">
          <span className="label">AI key (token, or NAME=value lines)</span>
          <textarea className="input h-20 py-2 font-mono text-[13px]" value={values.aiKey} onChange={(e) => set({ aiKey: e.target.value })} />
        </label>
        <label className="flex flex-col gap-1.5">
          <span className="label">Model</span>
          <select className="input" value={values.model} onChange={(e) => set({ model: e.target.value })}>
            {AI_MODELS.map((m) => (
              <option key={m}>{m}</option>
            ))}
          </select>
        </label>
      </div>
      <div className="flex items-center justify-end gap-3">
        {saved && <span className="text-[13px] text-green-700">Saved.</span>}
        <button type="button" className="btn-primary" onClick={() => void settings.set(values).then(() => setSaved(true))}>
          <SaveIcon /> Save
        </button>
      </div>
    </Card>
  );
}

/* --------------------------------- Config --------------------------------- */

const LABELS: Record<string, string> = {
  ...SCREEN_LABELS,
  aiPanel: "AI panel (every screen)",
  visualPreview: "Viewer › Guides: preview of the clicked visual",
  visualExplain: 'Viewer › Guides: "Explain" button',
  guideEditor: "Viewer › Guides: edit the guides in the viewer",
  excelUpload: "Excel editor: upload files into image / file fields",
  guides: "Guides",
  links: "Links",
  ai: "AI",
  visualPdf: "Visual as PDF (Guides)",
  visualPng: "Visual as PNG (Guides)",
  report: "Report export (PDF, .pbix)",
  all: "Every screen",
};

function Flags<T extends Record<string, boolean>>({ value, onChange }: { value: T; onChange: (v: T) => void }) {
  return (
    <div className="grid grid-cols-2 gap-x-6 gap-y-2">
      {Object.entries(value).map(([k, on]) => (
        <label key={k} className="flex items-center gap-2.5">
          <input type="checkbox" className="size-4 accent-[var(--color-primary)]" checked={on} onChange={(e) => onChange({ ...value, [k]: e.target.checked })} />
          {LABELS[k] ?? k}
        </label>
      ))}
    </div>
  );
}

/** The app's options for everyone: edited here, applied in this browser, shared through a JSON file on SharePoint. */
function ConfigCard() {
  const [draft, setDraft] = useState<HubConfig>(currentConfig);
  const [link, setLink] = useState("");
  const [state, setState] = useState<{ busy?: string; done?: string; error?: string }>({});
  useEffect(() => void settings.get().then((s) => setLink(s.configLink ?? "")), []);
  const set = <K extends keyof HubConfig>(key: K, value: HubConfig[K]) => setDraft((d) => ({ ...d, [key]: value }));
  const run = (busy: string, task: () => Promise<string>) => {
    setState({ busy });
    task().then((done) => setState({ done }), (e) => setState({ error: errorText(e) }));
  };
  const general = { aiPanel: draft.aiPanel, visualPreview: draft.visualPreview, visualExplain: draft.visualExplain, guideEditor: draft.guideEditor, excelUpload: draft.excelUpload };
  const section = (title: string, body: ReactNode) => (
    <div className="flex flex-col gap-2 border-t pt-4">
      <span className="font-medium">{title}</span>
      {body}
    </div>
  );
  return (
    <Card title="App options" hint="For everyone whose browser has the link of the options file (SharePoint / OneDrive). Apply them here, then save them to the file to share them.">
      <div className="flex gap-2">
        <input className="input" placeholder="Link of hub-config.json, or of a folder (the file is created in it)" value={link} onChange={(e) => setLink(e.target.value)} />
        <button
          type="button"
          className="btn"
          disabled={!link.trim() || !!state.busy}
          onClick={() =>
            run("Reading…", async () => {
              await loadHubConfig(link.trim());
              await settings.set({ configLink: link.trim() });
              setDraft(currentConfig());
              return "Options read from the file and applied.";
            })
          }
        >
          <DownloadIcon /> Load
        </button>
        <button
          type="button"
          className="btn"
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
        </button>
      </div>
      {state.busy && <p className="text-[13px] text-muted-foreground">{state.busy}</p>}
      {state.done && <p className="text-[13px] text-green-700">{state.done}</p>}
      {state.error && <p className="text-[13px] break-words text-destructive">{state.error}</p>}

      {section(
        "General",
        <>
          <div className="grid grid-cols-[1fr_auto] gap-3">
            <label className="flex flex-col gap-1.5">
              <span className="label">App name</span>
              <input className="input" value={draft.appTitle} onChange={(e) => set("appTitle", e.target.value)} />
            </label>
            <label className="flex flex-col gap-1.5">
              <span className="label">Primary color</span>
              <span className="flex items-center gap-2">
                <input type="color" className="h-9 w-11 rounded-lg border" value={/^#[0-9a-f]{6}$/i.test(draft.primaryColor) ? draft.primaryColor : "#18181b"} onChange={(e) => set("primaryColor", e.target.value)} />
                <input className="input w-28" value={draft.primaryColor} onChange={(e) => set("primaryColor", e.target.value)} />
              </span>
            </label>
          </div>
          <Flags value={general} onChange={(v) => setDraft((d) => ({ ...d, ...v }))} />
        </>,
      )}
      {section("Screens", <Flags value={draft.screens} onChange={(v) => set("screens", v)} />)}
      {section("Report viewer tabs", <Flags value={draft.viewerTabs} onChange={(v) => set("viewerTabs", v)} />)}
      {section("Exports", <Flags value={draft.export} onChange={(v) => set("export", v)} />)}
      {section(
        "AI instructions (company vocabulary, tone, rules)",
        (Object.keys(draft.prompts) as (keyof HubConfig["prompts"])[]).map((k) => (
          <label key={k} className="flex flex-col gap-1.5">
            <span className="label">{LABELS[k] ?? k}</span>
            <textarea className="input h-16 py-2" value={draft.prompts[k]} onChange={(e) => set("prompts", { ...draft.prompts, [k]: e.target.value })} />
          </label>
        )),
      )}
      <div className="sticky bottom-0 -mx-5 -mb-5 flex justify-end gap-2 rounded-b-xl border-t bg-card/95 px-5 py-3 backdrop-blur">
        <button type="button" className="btn" onClick={() => setDraft(structuredClone(DEFAULT_CONFIG))}>
          <RotateCcwIcon /> Defaults
        </button>
        <button type="button" className="btn-primary" onClick={() => (applyConfig(draft), setState({ done: "Applied in this browser (save to the file to share it)." }))}>
          <SaveIcon /> Apply here
        </button>
      </div>
    </Card>
  );
}
