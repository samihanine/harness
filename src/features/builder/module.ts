/** The Builder feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazyScreen as lazy } from "@/lib/lazy";
import type { Feature } from "@/lib/features";

export const builder: Feature = {
  id: "editor",
  path: "/editor",
  label: "Builder",
  description: "Build and change Power BI reports — pages, visuals, formatting — from your requests.",
  group: 0,
  picture: { outfit: "builder" },
  param: "report",
  collection: "reports",
  Screen: lazy(() => import("./screen")),
};
