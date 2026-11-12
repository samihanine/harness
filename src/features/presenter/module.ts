/** The Presenter feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazyScreen as lazy } from "@/lib/lazy";
import { PresentationIcon } from "lucide-react";
import { registerKind } from "@/lib/library";
import type { Feature } from "@/lib/features";

export const presenter: Feature = {
  id: "slides",
  path: "/slides",
  label: "Presenter",
  description: "Edit your PowerPoint decks with the assistant: texts, new slides from your template, Power BI visuals.",
  group: 0,
  picture: { outfit: "presenter" },
  param: "deck",
  collection: "presentations",
  Screen: lazy(() => import("./screen")),
  settings: [{ id: "slides", label: "Slides", icon: PresentationIcon, Panel: lazy(() => import("./settings")) }],
};

registerKind("presentations");
