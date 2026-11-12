/** The Excels feature: what the app needs to know of it (src/lib/features.ts); its screen loads when opened. */
import { lazyScreen as lazy } from "@/lib/lazy";
import { registerKind } from "@/lib/library";
import type { Feature } from "@/lib/features";

export const excel: Feature = {
  id: "excel",
  path: "/excel",
  label: "Excels",
  description: "Excel tables on SharePoint as simple databases: views, filters, columns, kept in sync with the file.",
  group: 1,
  picture: { logo: "excel" },
  param: "table",
  collection: "tables",
  linkable: { collection: "tables", title: "Excel tables" },
  Screen: lazy(() => import("./screen")),
};

registerKind("tables");
