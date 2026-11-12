/** The Reports feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazyScreen as lazy } from "@/lib/lazy";
import type { Feature } from "@/lib/features";

export const reports: Feature = {
  id: "reports",
  path: "/reports",
  label: "Reports",
  description: "Your Power BI reports: their pages, visuals, fields and notes.",
  group: 1,
  picture: { logo: "powerbi" },
  param: "report",
  collection: "reports",
  linkable: { collection: "reports", title: "Reports" },
  Screen: lazy(() => import("./screen")),
};
