import { useState } from "react";
import { Link, Outlet, createRootRoute } from "@tanstack/react-router";
import { BlocksIcon, ChartColumnIcon, CodeIcon, DatabaseIcon, MonitorPlayIcon, PanelLeftIcon, SearchCheckIcon, SettingsIcon } from "lucide-react";
import { SignInBanner } from "@/components/guards";
import { IconButton } from "@/components/icon-button";
import { SidePanel } from "@/components/side-panel";
import { TooltipProvider } from "@/components/ui/tooltip";

export const Route = createRootRoute({ component: Root });

const NAV = [
  { to: "/database", label: "Database", icon: DatabaseIcon },
  { to: "/dax", label: "DAX", icon: CodeIcon },
  { to: "/research", label: "Research", icon: SearchCheckIcon },
  { to: "/builder", label: "Builder", icon: BlocksIcon },
  { to: "/viewer", label: "Viewer", icon: MonitorPlayIcon },
] as const;

const nav =
  "flex items-center gap-1.5 rounded-md px-2 py-1 text-muted-foreground transition-colors hover:text-foreground data-[status=active]:bg-muted data-[status=active]:text-foreground [&_svg]:size-3.5";

function Root() {
  const [panel, setPanel] = useState(() => localStorage.getItem("pbi-agent:panel") !== "closed");
  const toggle = () => {
    localStorage.setItem("pbi-agent:panel", panel ? "closed" : "open");
    setPanel(!panel);
  };
  return (
    <TooltipProvider delay={300}>
      <div className="flex h-dvh flex-col">
        <header className="flex h-10 shrink-0 items-center gap-0.5 border-b px-2">
          <IconButton label={panel ? "Hide agent" : "Show agent"} onClick={toggle}>
            <PanelLeftIcon />
          </IconButton>
          <ChartColumnIcon className="mr-2 ml-1 size-4 text-muted-foreground" />
          {NAV.map(({ to, label, icon: Icon }) => (
            <Link key={to} to={to} className={nav}>
              <Icon /> {label}
            </Link>
          ))}
          <span className="ml-auto" />
          <Link to="/settings" className={nav} aria-label="Settings">
            <SettingsIcon />
          </Link>
        </header>
        <SignInBanner />
        <SidePanel open={panel}>
          <Outlet />
        </SidePanel>
      </div>
    </TooltipProvider>
  );
}
