/** The Confluence feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazy } from "react";
import type { Feature } from "@/lib/features";

export const confluence: Feature = {
  id: "confluence",
  path: "/confluence",
  label: "Confluence",
  description: "Confluence pages of a folder, edited as markdown with diagrams, and given to the AI as context.",
  group: 1,
  picture: { logo: "wiki" },
  param: "page",
  Screen: lazy(() => import("./screen")),
};
