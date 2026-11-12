/** The Data Scientist feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazyScreen as lazy } from "@/lib/lazy";
import { registerKind } from "@/lib/library";
import type { Feature } from "@/lib/features";

export const dataScientist: Feature = {
  id: "research",
  path: "/research",
  label: "Data Scientist",
  description: "Answer business questions from your models with figures, matrices and charts.",
  group: 0,
  picture: { outfit: "datascientist" },
  param: "research",
  collection: "researches",
  Screen: lazy(() => import("./screen")),
};

registerKind("researches");
