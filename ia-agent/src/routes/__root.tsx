import { Outlet, createRootRoute } from "@tanstack/react-router";
import { TooltipProvider } from "@/components/ui/tooltip";

export const Route = createRootRoute({
  component: () => (
    <TooltipProvider delay={300}>
      <Outlet />
    </TooltipProvider>
  ),
});
