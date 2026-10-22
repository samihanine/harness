/// <reference types="vite/client" />
import { useEffect, useState, type CSSProperties, type ReactNode } from "react";
import { HeadContent, Link, Navigate, Outlet, Scripts, createRootRoute, useLocation } from "@tanstack/react-router";
import { BookTextIcon, ChartPieIcon, DatabaseIcon, FileChartColumnIcon, FlaskConicalIcon, PencilRulerIcon, PresentationIcon, SettingsIcon, Table2Icon, type LucideIcon } from "lucide-react";
import { Toaster } from "@/components/ui/sonner";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { ensureBuiltIns } from "@/lib/library";
import { persist } from "@/lib/store";
import { applyTheme } from "@/lib/theme";
import { cn } from "@/lib/utils";
import { authStatus } from "@/server/ms";
import { loadHubConfig } from "@/lib/hub-config";
import { APP_TITLE, PRIMARY_COLOR, SCREENS, SCREEN_LABELS, SCREEN_PATHS, firstScreen, onConfigChange, type Screen } from "@/config";
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
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.6 ? "#0f172a" : "#ffffff";
}

function Shell({ children }: { children: ReactNode }) {
  return (
    <html lang="en" suppressHydrationWarning>
      <head>
        <HeadContent />
      </head>
      <body style={{ "--primary": PRIMARY_COLOR, "--ring": PRIMARY_COLOR, "--primary-foreground": readableOn(PRIMARY_COLOR) } as CSSProperties}>
        {children}
        <Scripts />
      </body>
    </html>
  );
}

const ICONS: Record<Screen, LucideIcon> = {
  viewer: ChartPieIcon,
  editor: PencilRulerIcon,
  research: FlaskConicalIcon,
  reports: FileChartColumnIcon,
  datasets: DatabaseIcon,
  excel: Table2Icon,
  contexts: BookTextIcon,
  slides: PresentationIcon,
};
/** Groups of the rail: work on reports, data sources, knowledge and output. */
const GROUPS: Screen[][] = [["viewer", "editor", "research"], ["reports", "datasets", "excel"], ["contexts", "slides"]];
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
    applyTheme();
    void loadHubConfig().catch((e) => console.warn("Config file not read", e));
    void ensureBuiltIns();
    persist();
  }, []);
  useEffect(() => {
    document.body.style.setProperty("--primary", PRIMARY_COLOR);
    document.body.style.setProperty("--ring", PRIMARY_COLOR);
    document.body.style.setProperty("--primary-foreground", readableOn(PRIMARY_COLOR));
  }, [version]);
  // A hidden screen is closed to direct links too; "/" opens the first screen.
  if (path === "/" || hiddenPaths().includes(path)) return <Navigate to={firstScreen()} search={{} as never} replace />;
  const item = (to: string, label: string, Icon: LucideIcon) => (
    <Tooltip key={to}>
      <TooltipTrigger asChild>
        <Link
          to={to}
          aria-label={label}
          className="group flex w-[60px] flex-col items-center gap-1 rounded-lg py-1.5 text-muted-foreground transition-colors hover:bg-accent hover:text-accent-foreground focus-visible:ring-[3px] focus-visible:ring-ring/30 outline-none"
          activeProps={{ className: "bg-accent !text-accent-foreground" }}
        >
          <Icon className="size-[19px]" strokeWidth={1.8} aria-hidden />
          <span className="text-[10.5px] leading-none font-medium">{label}</span>
        </Link>
      </TooltipTrigger>
      <TooltipContent side="right">{label}</TooltipContent>
    </Tooltip>
  );
  const groups = GROUPS.map((g) => g.filter((k) => SCREENS[k])).filter((g) => g.length);

  return (
    <TooltipProvider delayDuration={400}>
      <div className="flex h-full">
        <aside className="flex w-[72px] shrink-0 flex-col items-center border-r bg-sidebar py-3">
          <Link to={firstScreen()} aria-label={APP_TITLE} className="mb-3 flex size-9 items-center justify-center rounded-lg bg-primary text-[15px] font-bold text-primary-foreground shadow-sm">
            {APP_TITLE.slice(0, 1).toUpperCase()}
          </Link>
          <nav className="flex min-h-0 flex-1 flex-col items-center overflow-y-auto" aria-label="Screens">
            {groups.map((g, i) => (
              <div key={i} className={cn("flex flex-col items-center gap-0.5 py-2", i > 0 && "border-t")}>
                {g.map((k) => item(SCREEN_PATHS[k], SCREEN_LABELS[k], ICONS[k]))}
              </div>
            ))}
          </nav>
          <div className="flex flex-col items-center gap-2 border-t pt-2">
            {item("/settings", "Settings", SettingsIcon)}
            {status && !status.signedIn ? (
              <Link to="/settings" search={{ tab: "powerbi" } as never} aria-label="Sign in with Microsoft" className="flex size-8 items-center justify-center rounded-full bg-destructive text-[10.5px] font-semibold text-white">
                In
              </Link>
            ) : status?.account ? (
              <Tooltip>
                <TooltipTrigger asChild>
                  <span className="flex size-8 items-center justify-center rounded-full bg-primary/10 text-[12.5px] font-semibold text-primary">{status.account.slice(0, 1).toUpperCase()}</span>
                </TooltipTrigger>
                <TooltipContent side="right">{status.account}</TooltipContent>
              </Tooltip>
            ) : null}
          </div>
        </aside>
        <main key={version} className="min-w-0 flex-1">
          <Outlet />
        </main>
      </div>
      <Toaster position="bottom-right" richColors closeButton />
    </TooltipProvider>
  );
}
