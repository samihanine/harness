import { createHashHistory, createRouter } from "@tanstack/react-router";
import { routeTree } from "./routeTree.gen";

/** The standalone HTML file (export:html) is opened from the disk: its routes live in the URL hash. */
export function getRouter() {
  return createRouter({ routeTree, scrollRestoration: true, defaultPreload: "intent", ...(import.meta.env.MODE === "html" && { history: createHashHistory() }) });
}

declare module "@tanstack/react-router" {
  interface Register {
    router: ReturnType<typeof getRouter>;
  }
}
