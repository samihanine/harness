import { useState } from "react";
import { Link, Outlet, createRootRoute } from "@tanstack/react-router";
import { PanelRightIcon, TableIcon } from "lucide-react";
import { AgentFrame } from "@/components/agent-frame";
import { IconButton } from "@/components/icon-button";
import { TooltipProvider } from "@/components/ui/tooltip";

export const Route = createRootRoute({ component: Root });

const nav = "rounded-md px-2 py-1 text-muted-foreground transition-colors hover:text-foreground data-[status=active]:bg-muted data-[status=active]:text-foreground";

function Root() {
  const [agent, setAgent] = useState(() => localStorage.getItem("excel-agent:agent") !== "closed");
  const toggle = () => {
    localStorage.setItem("excel-agent:agent", agent ? "closed" : "open");
    setAgent(!agent);
  };
  return (
    <TooltipProvider delay={300}>
      <div className="flex h-dvh">
        <div className="flex min-w-0 flex-1 flex-col">
          <header className="flex h-10 shrink-0 items-center gap-1 border-b px-3">
            <TableIcon className="mr-1.5 size-4 text-muted-foreground" />
            <Link to="/files" className={nav}>
              Files
            </Link>
            <Link to="/schemas" className={nav}>
              Schemas
            </Link>
            <span className="ml-auto" />
            <IconButton label={agent ? "Hide agent" : "Show agent"} onClick={toggle}>
              <PanelRightIcon />
            </IconButton>
          </header>
          <main className="min-h-0 flex-1 overflow-auto">
            <Outlet />
          </main>
        </div>
        <AgentFrame open={agent} />
      </div>
    </TooltipProvider>
  );
}
