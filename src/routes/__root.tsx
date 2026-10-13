/// <reference types="vite/client" />
import { useEffect, useState, type CSSProperties, type ReactNode } from "react";
import { HeadContent, Link, Navigate, Outlet, Scripts, createRootRoute, useLocation } from "@tanstack/react-router";
import { ChevronDownIcon, ChevronUpIcon, SettingsIcon } from "lucide-react";
import { last } from "@/lib/last";
import { authStatus } from "@/server/ms";
import { loadHubConfig } from "@/lib/hub-config";
import { APP_TITLE, PRIMARY_COLOR, SCREENS, SCREEN_PATHS, SHOW_HEADER, firstScreen, onConfigChange } from "@/config";
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
export function readableOn(color: string) {
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

const nav = () =>
  (
    [
      [SCREEN_PATHS.viewer, "Report Viewer", SCREENS.viewer],
      [SCREEN_PATHS.builder, "Report builder", SCREENS.builder],
      [SCREEN_PATHS.datasets, "Dataset builder", SCREENS.datasets],
      [SCREEN_PATHS.dax, "DAX builder", SCREENS.dax],
      [SCREEN_PATHS.pivot, "TCD builder", SCREENS.pivot],
      [SCREEN_PATHS.library, "Library", SCREENS.library],
    ] as const
  ).filter(([, , shown]) => shown);
const hiddenPaths = () => (Object.keys(SCREENS) as (keyof typeof SCREENS)[]).filter((k) => !SCREENS[k]).map((k) => SCREEN_PATHS[k]);

type Status = Awaited<ReturnType<typeof authStatus>>;

function Layout() {
  const [status, setStatus] = useState<Status>();
  const path = useLocation({ select: (l) => l.pathname });
  // Read again on each screen change (e.g. after signing in from Settings).
  useEffect(() => void authStatus().then(setStatus, () => setStatus(undefined)), [path]);
  // Options: the SharePoint file is read at start; any change draws the screens again.
  const [version, setVersion] = useState(0);
  useEffect(() => onConfigChange(() => setVersion((v) => v + 1)), []);
  useEffect(() => void loadHubConfig().catch((e) => console.warn("Config file not read", e)), []);
  useEffect(() => document.body.style.setProperty("--color-primary-foreground", readableOn(PRIMARY_COLOR)), [version]);
  // The user can fold the header (remembered in this browser); a small button brings it back.
  const [folded, setFolded] = useState(() => last.get("header") === "folded");
  const fold = (next: boolean) => {
    last.set("header", next ? "folded" : "open");
    setFolded(next);
  };
  // A hidden screen is closed to direct links too.
  if (hiddenPaths().includes(path)) return <Navigate to={firstScreen()} search={{} as never} replace />;
  // Hidden header: still shown while the user has to sign in.
  const mustSignIn = status !== undefined && !status.signedIn;
  const header = (SHOW_HEADER && !folded) || mustSignIn;

  return (
    <div className="flex h-full flex-col">
      {header && (
        <header className="flex h-11 shrink-0 items-center gap-1 border-b bg-card px-3">
          <span className="mr-3 font-semibold" style={{ color: PRIMARY_COLOR }}>
            {APP_TITLE}
          </span>
          {nav().map(([to, label]) => (
            <Link key={to} to={to} className="rounded-md px-2.5 py-1 text-muted-foreground hover:text-foreground" activeProps={{ className: "bg-muted !text-foreground" }} activeOptions={{ exact: to === "/" }}>
              {label}
            </Link>
          ))}
          <div className="ml-auto flex items-center gap-2 text-[12px] text-muted-foreground">
            {status?.signedIn ? (
              status.account
            ) : status ? (
              <Link to="/settings" search={{ tab: "microsoft" }} className="btn-primary h-7">
                Sign in with Microsoft
              </Link>
            ) : null}
            <Link to="/settings" search={{ tab: undefined }} className="icon-btn" title="Settings (config, AI, Microsoft account)" activeProps={{ className: "bg-muted" }}>
              <SettingsIcon />
            </Link>
            {SHOW_HEADER && !mustSignIn && (
              <button type="button" className="icon-btn" title="Hide the top bar" onClick={() => fold(true)}>
                <ChevronUpIcon />
              </button>
            )}
          </div>
        </header>
      )}
      {SHOW_HEADER && folded && !header && (
        <button type="button" title="Show the top bar" onClick={() => fold(false)} className="fixed top-0 left-1/2 z-50 flex h-4 w-10 -translate-x-1/2 items-center justify-center rounded-b-md border border-t-0 bg-card text-muted-foreground shadow-sm hover:h-5 hover:text-foreground [&_svg]:size-3.5">
          <ChevronDownIcon />
        </button>
      )}
      <main key={version} className="min-h-0 flex-1">
        <Outlet />
      </main>
    </div>
  );
}
