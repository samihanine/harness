/** The Datasets feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazyScreen as lazy } from "@/lib/lazy";
import type { Feature } from "@/lib/features";

export const datasets: Feature = {
  id: "datasets",
  path: "/datasets",
  label: "Datasets",
  description: "Your semantic models: tables, measures, relationships, DAX and Power Query.",
  group: 1,
  picture: { logo: "model" },
  param: "dataset",
  collection: "datasets",
  linkable: { collection: "datasets", title: "Datasets" },
  Screen: lazy(() => import("./screen")),
};
