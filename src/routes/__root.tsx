/// <reference types="vite/client" />
import { useEffect, useState, type CSSProperties, type ReactNode } from "react";
import { HeadContent, Link, Outlet, Scripts, createRootRoute } from "@tanstack/react-router";
import { LogOutIcon, SettingsIcon } from "lucide-react";
import { authStatus, signIn, signOut } from "@/server/ms";
import { AI_MODELS } from "@/agent/llm";
import { settings } from "@/lib/store";
import { APP_TITLE, PRIMARY_COLOR, SCREENS } from "@/config";
import css from "@/styles.css?url";

export const Route = createRootRoute({
  head: () => ({
    meta: [{ charSet: "utf-8" }, { name: "viewport", content: "width=device-width, initial-scale=1" }, { title: APP_TITLE }],
    links: [{ rel: "stylesheet", href: css }],
  }),
  shellComponent: Shell,
  component: Layout,
});

/** Black or white text, whichever reads best on a color (#rrggbb; anything else: white). */
function readableOn(color: string) {
  const m = color.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return "#ffffff";
  const [r, g, b] = m.slice(1).map((h) => parseInt(h, 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.6 ? "#18181b" : "#ffffff";
}

function Shell({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <head>
        <HeadContent />
      </head>
      <body style={{ "--color-primary": PRIMARY_COLOR, "--color-primary-foreground": readableOn(PRIMARY_COLOR) } as CSSProperties}>
        {children}
        <Scripts />
      </body>
    </html>
  );
}

const NAV = (
  [
    ["/viewer", "Report Viewer", SCREENS.viewer],
    ["/builder", "Report builder", SCREENS.builder],
    ["/datasets", "Dataset builder", SCREENS.datasets],
    ["/", "Library", SCREENS.library],
  ] as const
).filter(([, , shown]) => shown);

type Status = Awaited<ReturnType<typeof authStatus>>;

function Layout() {
  const [status, setStatus] = useState<Status>();
  const [open, setOpen] = useState(false);
  useEffect(() => {
    let stop = false;
    const poll = async () => {
      const s = await authStatus().catch(() => undefined);
      if (!stop) setStatus(s);
      if (!stop && s?.pending) setTimeout(poll, 3000);
    };
    void poll();
    return () => void (stop = true);
  }, [status?.pending?.userCode]);

  return (
    <div className="flex h-full flex-col">
      <header className="flex h-11 shrink-0 items-center gap-1 border-b bg-card px-3">
        <span className="mr-3 font-semibold" style={{ color: PRIMARY_COLOR }}>
          {APP_TITLE}
        </span>
        {NAV.map(([to, label]) => (
          <Link key={to} to={to} className="rounded-md px-2.5 py-1 text-muted-foreground hover:text-foreground" activeProps={{ className: "bg-muted !text-foreground" }} activeOptions={{ exact: to === "/" }}>
            {label}
          </Link>
        ))}
        <div className="ml-auto flex items-center gap-2 text-[12px] text-muted-foreground">
          {status?.signedIn ? (
            <>
              {status.account}
              <button type="button" className="icon-btn" title="Sign out" onClick={() => void signOut().then(() => setStatus(undefined))}>
                <LogOutIcon />
              </button>
            </>
          ) : status?.pending ? (
            <span>
              Open <a className="underline" href={status.pending.verificationUri} target="_blank" rel="noreferrer">{status.pending.verificationUri}</a> and enter <b className="text-foreground">{status.pending.userCode}</b>
              {status.pending.error && <span className="text-destructive"> — {status.pending.error}</span>}
            </span>
          ) : (
            <button type="button" className="btn-primary h-7" onClick={() => void signIn().then((pending) => setStatus({ signedIn: false, pending, account: undefined }))}>
              Sign in with Microsoft
            </button>
          )}
          <button type="button" className="icon-btn" title="AI settings" onClick={() => setOpen(!open)}>
            <SettingsIcon />
          </button>
        </div>
      </header>
      {open && <AiSettings onClose={() => setOpen(false)} />}
      <main className="min-h-0 flex-1">
        <Outlet />
      </main>
    </div>
  );
}

function AiSettings({ onClose }: { onClose: () => void }) {
  const [values, setValues] = useState({ aiKey: "", model: AI_MODELS[0] as string });
  useEffect(() => void settings.get().then((s) => setValues({ aiKey: s.aiKey ?? "", model: s.model ?? AI_MODELS[0] })), []);
  return (
    <div className="absolute top-12 right-3 z-50 flex w-80 flex-col gap-2 rounded-lg border bg-card p-3 shadow-lg">
      <label className="label">AI key (token, or NAME=value lines)</label>
      <textarea className="input h-16 py-1 font-mono" value={values.aiKey} onChange={(e) => setValues({ ...values, aiKey: e.target.value })} />
      <label className="label">Model</label>
      <select className="input" value={values.model} onChange={(e) => setValues({ ...values, model: e.target.value })}>
        {AI_MODELS.map((m) => (
          <option key={m}>{m}</option>
        ))}
      </select>
      <button type="button" className="btn-primary self-end" onClick={() => void settings.set(values).then(onClose)}>
        Save
      </button>
    </div>
  );
}
