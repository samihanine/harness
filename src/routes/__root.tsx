/// <reference types="vite/client" />
import { useEffect, useState, type ReactNode } from "react";
import { Link, Navigate, Outlet, createRootRoute, useLocation } from "@tanstack/react-router";
import { SettingsIcon } from "lucide-react";
import { ScreenLogo } from "@/components/screen-logo";
import { HealthBanner } from "@/components/health-banner";
import { Toaster } from "@/components/ui/sonner";
import { TooltipProvider } from "@/components/ui/tooltip";
import { ensureBuiltIns } from "@/lib/library";
import { persist } from "@/lib/store";
import { applyTheme } from "@/lib/theme";
import { cn } from "@/lib/utils";
import { tokenStatus } from "@/services/ms";
import { loadHubConfig } from "@/services/hub-config";
import { APP_TITLE, HOME, PRIMARY_COLOR, SCREENS, SCREEN_GROUPS, SCREEN_LABELS, SCREEN_PATHS, firstScreen, onConfigChange, type Screen } from "@/config";
import "@/styles.css";

/** The app is a single page: index.html loads src/main.tsx, which mounts the router (no server). */
export const Route = createRootRoute({ component: Layout });

/** Black or white text, whichever reads best on a color (#rrggbb; anything else: white). */
export function readableOn(color: string) {
  const m = color.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return "#ffffff";
  const [r, g, b] = m.slice(1).map((h) => parseInt(h, 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.6 ? "#0f172a" : "#ffffff";
}

/**
 * The colors made from the app's primary color (Settings): buttons, focus rings, selected items. The
 * theme's --color-* are set too: on <html> they were resolved with the default blue.
 */
const colorVars = () => {
  // A see-through tint (no other variable in it: it changes at once, in light and dark themes).
  const accent = `color-mix(in srgb, ${PRIMARY_COLOR} 12%, transparent)`;
  const accentText = PRIMARY_COLOR;
  return {
    "--primary": PRIMARY_COLOR,
    "--color-primary": PRIMARY_COLOR,
    "--ring": PRIMARY_COLOR,
    "--color-ring": PRIMARY_COLOR,
    "--primary-foreground": readableOn(PRIMARY_COLOR),
    "--color-primary-foreground": readableOn(PRIMARY_COLOR),
    "--accent": accent,
    "--color-accent": accent,
    "--accent-foreground": accentText,
    "--color-accent-foreground": accentText,
  };
};


const hiddenPaths = () => (Object.keys(SCREENS) as Screen[]).filter((k) => !SCREENS[k]).map((k) => SCREEN_PATHS[k]);

type Status = Awaited<ReturnType<typeof tokenStatus>>;

function Layout() {
  const [status, setStatus] = useState<Status>();
  const path = useLocation({ select: (l) => l.pathname });
  // Read again on each screen change (e.g. after signing in from Settings).
  useEffect(() => void tokenStatus().then(setStatus, () => setStatus(undefined)), [path]);
  // Options: the SharePoint file is read at start; any change draws the screens again.
  const [version, setVersion] = useState(0);
  useEffect(() => onConfigChange(() => setVersion((v) => v + 1)), []);
  useEffect(() => {
    applyTheme();
    void loadHubConfig().catch((e) => console.warn("Config file not read", e));
    void ensureBuiltIns();
    persist();
  }, []);
  useEffect(() => {
    // The tints made from the primary follow it too (the selected item of the sidebar and the lists).
    for (const [k, v] of Object.entries(colorVars())) document.body.style.setProperty(k, v);
  }, [version]);
  // A hidden screen is closed to direct links too; "/" opens the first screen when there is no home page.
  if ((path === "/" && !HOME) || hiddenPaths().includes(path)) return <Navigate to={firstScreen()} search={{} as never} replace />;
  /** A screen in the sidebar: its picture, its name under it. */
  const item = (to: string, label: string, picture: ReactNode) => (
    <Link
      key={to}
      to={to}
      className="flex w-full flex-col items-center gap-1 rounded-lg px-1 pt-2 pb-1.5 text-[11.5px] font-medium text-muted-foreground transition-colors outline-none hover:bg-accent hover:text-accent-foreground focus-visible:ring-[3px] focus-visible:ring-ring/30"
      activeProps={{ className: "bg-accent !text-accent-foreground" }}
    >
      <span className="flex h-10 items-center justify-center">{picture}</span>
      <span className="line-clamp-2 max-w-full text-center leading-tight">{label}</span>
    </Link>
  );
  const groups = SCREEN_GROUPS.map((g) => g.filter((k) => SCREENS[k])).filter((g) => g.length);

  return (
    <TooltipProvider delayDuration={400}>
      <div className="flex h-full overflow-hidden">
        <aside className="flex w-[92px] shrink-0 flex-col border-r bg-sidebar px-1.5 py-3">
          <Link to={HOME ? "/" : firstScreen()} aria-label={APP_TITLE} className="mx-auto mb-2 flex size-9 items-center justify-center rounded-lg bg-primary text-[15px] font-bold text-primary-foreground shadow-sm outline-none focus-visible:ring-[3px] focus-visible:ring-ring/30">
            {APP_TITLE.slice(0, 1).toUpperCase()}
          </Link>
          <nav className="flex min-h-0 flex-1 flex-col overflow-y-auto" aria-label="Screens">
            {groups.map((g, i) => (
              <div key={i} className={cn("flex flex-col gap-0.5 py-2", i > 0 && "border-t")}>
                {g.map((k) => item(SCREEN_PATHS[k], SCREEN_LABELS[k], <ScreenLogo screen={k} size={k in { reports: 1, datasets: 1, excel: 1, contexts: 1, confluence: 1 } ? 30 : 38} />))}
              </div>
            ))}
          </nav>
          <div className="flex flex-col border-t pt-2">{item("/settings", "Settings", <SettingsIcon className="size-6" strokeWidth={1.7} aria-hidden />)}</div>
        </aside>
        <main key={version} className="flex min-w-0 flex-1 flex-col">
          <HealthBanner signedIn={status?.signedIn} />
          <div className="min-h-0 flex-1">
            <Outlet />
          </div>
        </main>
      </div>
      <Toaster position="bottom-right" richColors closeButton />
    </TooltipProvider>
  );
}
