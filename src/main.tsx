/** Entry of the single-page app: the router mounted in #root (hash routes in the standalone HTML export). */
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { RouterProvider } from "@tanstack/react-router";
import { getRouter } from "./router";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <RouterProvider router={getRouter()} />
  </StrictMode>,
);
