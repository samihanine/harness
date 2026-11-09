/**
 * The app layer: the only place that knows every feature. It registers them (sidebar order) for the
 * shared code; each route file (src/routes) hosts one of them. Adding a screen: its folder in
 * src/features with a module.ts, a line here, a route file, and its flag in config.ts (SCREENS).
 */
import { registerFeatures } from "@/lib/features";
import { trainer } from "@/features/trainer/module";
import { builder } from "@/features/builder/module";
import { dataScientist } from "@/features/data-scientist/module";
import { presenter } from "@/features/presenter/module";
import { reports } from "@/features/reports/module";
import { datasets } from "@/features/datasets/module";
import { excel } from "@/features/excel/module";
import { confluence } from "@/features/confluence/module";

registerFeatures([trainer, builder, dataScientist, presenter, reports, datasets, excel, confluence]);
