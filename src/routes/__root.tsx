/// <reference types="vite/client" />
import { useEffect, useState, type CSSProperties, type ReactNode } from "react";
import { HeadContent, Link, Navigate, Outlet, Scripts, createRootRoute, useLocation } from "@tanstack/react-router";
import { ChartPieIcon, ChevronLeftIcon, ChevronRightIcon, DatabaseIcon, FlaskConicalIcon, LibraryIcon, PencilRulerIcon, SettingsIcon, ShapesIcon, Table2Icon, type LucideIcon } from "lucide-react";
import { last } from "@/lib/last";
import { ensureBuiltIns } from "@/lib/library";
import { persist } from "@/lib/store";
import { authStatus } from "@/server/ms";
import { loadHubConfig } from "@/lib/hub-config";
import { APP_TITLE, PRIMARY_COLOR, SCREENS, SCREEN_PATHS, firstScreen, onConfigChange, type Screen } from "@/config";
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
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.6 ? "#111827" : "#ffffff";
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

/** Labels under the icons of the rail. */
const SHORT: Record<Screen, string> = { viewer: "Viewer", editor: "Editor", model: "Dataset", research: "Research", schemas: "Schema", excel: "Excel", library: "Library" };
const ICONS: Record<Screen, LucideIcon> = { viewer: ChartPieIcon, editor: PencilRulerIcon, model: DatabaseIcon, research: FlaskConicalIcon, schemas: ShapesIcon, excel: Table2Icon, library: LibraryIcon };
/** Screens in the menu, in the order of SCREENS (read at render: an admin may hide some). */
const nav = () => (Object.keys(SCREENS) as Screen[]).filter((k) => SCREENS[k]);
const hiddenPaths = () => (Object.keys(SCREENS) as Screen[]).filter((k) => !SCREENS[k]).map((k) => SCREEN_PATHS[k]);

type Status = Awaited<ReturnType<typeof authStatus>>;

function Layout() {
  const [status, setStatus] = useState<Status>();
  const path = useLocation({ select: (l) => l.pathname });
  // Read again on each screen change (e.g. after signing in from Settings).
  useEffect(() => void authStatus().then(setStatus, () => setStatus(undefined)), [path]);
  // Options: the SharePoint file is read at start; any change draws the screens again.
  const [version, setVersion] = useState(0);
  useEffect(() => onConfigChange(() => setVersion((v) => v + 1)), []);
  useEffect(() => {
    void loadHubConfig().catch((e) => console.warn("Config file not read", e));
    void ensureBuiltIns();
    persist();
  }, []);
  useEffect(() => document.body.style.setProperty("--color-primary-foreground", readableOn(PRIMARY_COLOR)), [version]);
  // The menu can be hidden (remembered in this browser); a small tab brings it back.
  const [folded, setFolded] = useState(() => last.get("sidebar") === "folded");
  const fold = (next: boolean) => {
    last.set("sidebar", next ? "folded" : "open");
    setFolded(next);
  };
  // A hidden screen is closed to direct links too.
  if (hiddenPaths().includes(path)) return <Navigate to={firstScreen()} search={{} as never} replace />;
  const item = (to: string, label: string, Icon: LucideIcon, exact = false) => (
    <Link
      key={to}
      to={to}
      title={label}
      className="flex w-[60px] flex-col items-center gap-1 rounded-xl py-1.5 text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
      activeProps={{ className: "bg-muted !text-foreground" }}
      activeOptions={{ exact }}
    >
      <Icon className="size-5" strokeWidth={1.7} />
      <span className="text-[10.5px] leading-none font-medium">{label}</span>
    </Link>
  );
  const shown = nav();
  /** Groups of the rail: reports, data, schemas, library. */
  const groups = [["viewer", "editor"], ["model", "research"], ["schemas", "excel"], ["library"]].map((g) => g.filter((k) => shown.includes(k as Screen)) as Screen[]).filter((g) => g.length);

  return (
    <div className="flex h-full">
      {folded ? (
        <button type="button" title="Show the menu" onClick={() => fold(false)} className="fixed bottom-3 left-0 z-50 flex h-9 w-4 items-center justify-center rounded-r-md border border-l-0 bg-card text-muted-foreground shadow-sm hover:w-5 hover:text-foreground">
          <ChevronRightIcon className="size-3.5" />
        </button>
      ) : (
        <aside className="flex w-[72px] shrink-0 flex-col items-center border-r bg-card py-3">
          <nav className="flex min-h-0 flex-1 flex-col items-center overflow-y-auto">
            {groups.map((g, i) => (
              <div key={i} className="flex flex-col items-center gap-0.5 border-t py-2 first:border-t-0 first:pt-0">
                {g.map((k) => item(SCREEN_PATHS[k], SHORT[k], ICONS[k], k === "library"))}
              </div>
            ))}
          </nav>
          <div className="flex flex-col items-center gap-1.5 border-t pt-2">
            {item("/settings", "Settings", SettingsIcon)}
            {status && !status.signedIn ? (
              <Link to="/settings" title="Sign in with Microsoft" className="flex size-9 items-center justify-center rounded-full bg-destructive text-[11px] font-semibold text-white">
                In
              </Link>
            ) : status?.account ? (
              <Link to="/" title={`${status.account} — home`} className="flex size-9 items-center justify-center rounded-full bg-muted text-[13px] font-semibold hover:bg-zinc-200">
                {status.account.slice(0, 1).toUpperCase()}
              </Link>
            ) : null}
            <button type="button" className="icon-btn size-7" title="Hide the menu" onClick={() => fold(true)}>
              <ChevronLeftIcon />
            </button>
          </div>
        </aside>
      )}
      <main key={version} className="min-w-0 flex-1">
        <Outlet />
      </main>
    </div>
  );
}
