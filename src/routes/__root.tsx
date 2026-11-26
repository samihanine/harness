/// <reference types="vite/client" />
import { useEffect, useState } from "react";
import { Link, Outlet, createRootRoute, useLocation } from "@tanstack/react-router";
import { HomeIcon, SettingsIcon } from "lucide-react";
import { ScreenIcon } from "@/components/screen-icons";
import { Toaster } from "@/components/ui/sonner";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { SCREEN_INFO } from "@/components/screens";
import { SCREENS } from "@shared/types";
import { useSettings } from "@/lib/settings";
import { applyTheme } from "@/lib/theme";
import { cn } from "@/lib/utils";
import { get } from "@/lib/api";
import "@/styles.css";

export const Route = createRootRoute({ component: Layout });

/** Black or white text on a color. */
export function readableOn(color: string) {
  const m = color.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!m) return "#ffffff";
  const [r, g, b] = m.slice(1).map((h) => parseInt(h, 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.6 ? "#111111" : "#ffffff";
}

function RailLink({ to, label, children, active }: { to: string; label: string; children: React.ReactNode; active: boolean }) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <Link to={to} aria-label={label} className={cn("group flex size-12 items-center justify-center rounded-2xl text-muted-foreground transition-all", active ? "border bg-card text-foreground shadow-[0_2px_10px_rgba(0,0,0,0.06)]" : "hover:bg-muted hover:text-foreground")}>
          {children}
        </Link>
      </TooltipTrigger>
      <TooltipContent side="right">{label}</TooltipContent>
    </Tooltip>
  );
}

function Layout() {
  const settings = useSettings();
  const path = useLocation({ select: (l) => l.pathname });
  const [signedIn, setSignedIn] = useState<boolean>();
  useEffect(() => applyTheme(), []);
  useEffect(() => void get("/auth").then((s) => setSignedIn(s.signedIn), () => setSignedIn(false)), [path]);
  const primary = settings?.general.primaryColor ?? "#b39251";
  const style = { "--primary": primary, "--color-primary": primary, "--ring": primary, "--color-ring": primary, "--primary-foreground": readableOn(primary), "--color-primary-foreground": readableOn(primary) } as React.CSSProperties;
  const shown = SCREENS.filter((s) => settings?.general.screens[s] !== false);
  return (
    <TooltipProvider delayDuration={300}>
      <div style={style} className="h-screen bg-app p-2.5 sm:p-3">
        <div className="flex h-full overflow-hidden rounded-[26px] border bg-sidebar shadow-[0_1px_2px_rgba(0,0,0,0.04)]">
          <nav className="flex w-[76px] shrink-0 flex-col items-center gap-2 py-5">
            <RailLink to="/" label="Home" active={path === "/"}>
              <HomeIcon className="size-5" />
            </RailLink>
            <div className="mb-2 h-px w-8 bg-border" />
            {shown.map((s) => {
              const info = SCREEN_INFO[s];
              return (
                <RailLink key={s} to={info.path} label={info.label} active={path.startsWith(info.path)}>
                  <ScreenIcon screen={s} />
                </RailLink>
              );
            })}
            <div className="flex-1" />
            <RailLink to="/settings" label={signedIn === false ? "Settings — sign in to Microsoft" : "Settings"} active={path.startsWith("/settings")}>
              <span className="relative">
                <SettingsIcon className="size-[22px] transition-transform duration-500 group-hover:rotate-90" />
                {signedIn === false && <span className="absolute -top-0.5 -right-0.5 size-2.5 rounded-full bg-primary ring-2 ring-sidebar" />}
              </span>
            </RailLink>
          </nav>
          <main className="min-w-0 flex-1 bg-background p-2.5 sm:p-3">
            <Outlet />
          </main>
        </div>
        <Toaster position="bottom-right" />
      </div>
    </TooltipProvider>
  );
}
