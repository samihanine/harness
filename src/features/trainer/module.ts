/** The Trainer feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazy } from "react";
import { BookOpenIcon } from "lucide-react";
import { onReportAdded } from "@/lib/library";
import type { Feature } from "@/lib/features";

export const trainer: Feature = {
  id: "viewer",
  path: "/viewer",
  label: "Trainer",
  description: "Read a Power BI report with an assistant that explains it, filters it and answers from its data.",
  group: 0,
  picture: { outfit: "trainer" },
  param: "report",
  collection: "reports",
  Screen: lazy(() => import("./screen")),
  settings: [{ id: "guides", label: "Guides", icon: BookOpenIcon, Panel: lazy(() => import("./settings")) }],
};

// Adding a report makes an empty guide for each of its data visuals (the downloaded file lists them all).
onReportAdded(async (entry, schema) => (await import("./guides")).addEmptyGuides(entry.id, schema));
