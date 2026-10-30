/**
 * Live scenarios (EVAL_LIVE=1): the agents work on the real test model of sales.ts (My workspace)
 * through the real Microsoft APIs; the checks read back what was written (canvas, report definitions).
 */
import type { Agent } from "@/agent/loop";
import { builderAgent } from "@/features/builder/builder";
import { datasetText, withContexts } from "@/agent/context";
import { modelTools } from "@/agent/model-tools";
import { DAX_RULES, daxTool } from "@/agent/pbi-tools";
import { researchAgent } from "@/features/data-scientist/research-tools";
import { addDataset } from "@/lib/library";
import { definition, partText } from "@/services/report-def";
import { newResearch, readResearch, type ResearchEntry } from "@/features/data-scientist/research";
import { createBlankReport } from "@/features/builder/new-report";
import * as store from "@/lib/store";
import type { DatasetEntry, ReportEntry } from "@/lib/store";
import { FIXTURES, readFixtures } from "./sales";
import type { Scenario } from "./scenarios";
import { writeFileSync } from "node:fs";

const { peek, reset } = store as unknown as typeof import("./fake-store");
const expect = (ok: unknown, problem: string) => (ok ? [] : [problem]);

const fixtures = () => {
  const f = readFixtures();
  if (!f) throw new Error("No test model: run `bun run eval:live:setup` first.");
  return f;
};

/** The test model as the app knows it (structure read from its TMDL), read once per run. */
let dataset: DatasetEntry | undefined;
async function testDataset() {
  if (dataset) return dataset;
  const f = fixtures();
  reset({});
  dataset = await addDataset(f.datasetId, undefined, f.datasetName);
  if (!dataset.info) throw new Error(`The test model's structure could not be read: ${dataset.error}`);
  return dataset;
}

/** The Builder's report on the test model: made once (a first matrix page), kept in the fixtures. */
async function testReport(d: DatasetEntry): Promise<ReportEntry> {
  const f = fixtures();
  if (!f.reportId) {
    const name = "Hub test — Builder (temporaire)";
    const id = await createBlankReport(name, { id: d.id, groupId: d.groupId });
    Object.assign(f, { reportId: id, reportName: name });
    writeFileSync(FIXTURES, JSON.stringify(f, null, 2));
  }
  return { id: f.reportId!, name: f.reportName!, url: "", datasetId: d.id, access: "edit", contextIds: [] };
}

/** Visuals of a page of a report definition (type, position, fields, filters). */
async function pageVisuals(reportId: string, match: RegExp) {
  const parts = await definition({ id: reportId });
  const json = (p: { payload: string }) => JSON.parse(partText(p as never));
  const page = parts.filter((p) => /^definition\/pages\/[^/]+\/page\.json$/.test(p.path)).map((p) => ({ path: p.path, ...json(p) })).find((p) => match.test(p.displayName));
  if (!page) return { page: undefined, visuals: [], pages: parts.filter((p) => p.path.endsWith("/page.json")).map((p) => json(p).displayName as string) };
  const visuals = parts.filter((p) => p.path.startsWith(`definition/pages/${page.name}/visuals/`) && p.path.endsWith("visual.json")).map((p) => json(p));
  return { page, visuals, pages: [] as string[] };
}

/** Overlapping or out-of-page visuals (the check a careful designer would do). */
function layoutProblems(visuals: any[], width: number, height: number) {
  const boxes = visuals.filter((v) => v.position).map((v) => ({ name: `${v.visual?.visualType ?? "group"} ${v.name.slice(0, 6)}`, ...v.position }));
  const out = boxes.filter((b) => b.x < 0 || b.y < 0 || b.x + b.width > width + 1 || b.y + b.height > height + 1).map((b) => `${b.name} outside the page`);
  const overlaps: string[] = [];
  for (let i = 0; i < boxes.length; i++)
    for (let j = i + 1; j < boxes.length; j++) {
      const a = boxes[i], b = boxes[j];
      const w = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x);
      const h = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y);
      if (w > 4 && h > 4) overlaps.push(`${a.name} overlaps ${b.name}`);
    }
  return [...out, ...overlaps];
}

const fieldsOf = (v: any) => JSON.stringify(v.visual?.query?.queryState ?? {});

export const LIVE_SCENARIOS: Scenario[] = [
  {
    name: "research-sales",
    description: "Data Scientist (live): growth by category, margin by channel — the answer in the canvas text, pivots justifying it",
    request: "Quelles catégories de produits progressent le plus en 2025 par rapport à 2024 ? Et comment évolue la marge par canal ?",
    setup: async (): Promise<Agent> => {
      const d = await testDataset();
      reset({ datasets: [d], researches: [{ ...newResearch("Test ventes"), id: "r1", datasetIds: [d.id] }], reports: [], contexts: [] });
      return researchAgent(() => readResearch(peek("researches")[0]), () => peek("datasets"));
    },
    check: async () => {
      const r = readResearch(peek("researches")[0]) as ResearchEntry;
      return [
        ...expect(r.text.length > 80, "no answer in the canvas text"),
        ...expect(/cat[ée]gor/i.test(r.text) && /marge|canal/i.test(r.text), "the text does not answer both questions"),
        ...expect(r.cards.some((c) => c.kind === "pivot"), "no pivot justifying the answer"),
        ...r.cards.flatMap((c) => (c.error || !c.rows?.length ? [`card "${c.title}": ${c.error ?? "no rows"}`] : [])),
      ];
    },
  },
  {
    name: "research-pasted",
    description: "Data Scientist (live): a table pasted from a report (its labels are not model names) and a check of the user's reading — the 29/10 failure",
    request:
      "Catégorie\tVentes 2025\tvs Objectif %\nBikes\t2,4M\t+12%\nAccessories\t310K\t-4%\n\nDe ce que je comprends, la colonne « vs Objectif % » est l'écart en pourcentage entre les ventes et l'objectif. C'est bien ça ? Vérifie avec les données.",
    setup: async (): Promise<Agent> => {
      const d = await testDataset();
      reset({ datasets: [d], researches: [{ ...newResearch("Lecture"), id: "r3", datasetIds: [d.id] }], reports: [], contexts: [] });
      return researchAgent(() => readResearch(peek("researches")[0]), () => peek("datasets"));
    },
    check: async () => {
      const r = readResearch(peek("researches")[0]) as ResearchEntry;
      return [
        ...expect(r.text.length > 80, "no answer in the canvas text"),
        ...expect(r.cards.some((c) => c.kind === "pivot"), "no pivot checking the figures"),
        ...r.cards.flatMap((c) => (c.error || !c.rows?.length ? [`card "${c.title}": ${c.error ?? "no rows"}`] : [])),
      ];
    },
  },
  {
    name: "builder-overview",
    description: "Builder (live): a full overview page (KPIs, trend vs LY, top 10, channel donut, slicer)",
    request:
      "Crée une page « Vue d'ensemble » : 4 cartes KPI (ventes, marge %, commandes, évolution des ventes vs N-1), une courbe des ventes mensuelles comparées à l'an dernier, le top 10 des produits par ventes, la répartition des ventes par canal en anneau et un segment pour choisir l'année. Fais une mise en page propre et cohérente.",
    setup: async () => {
      const d = await testDataset();
      const entry = await testReport(d);
      reset({ datasets: [d], reports: [entry], contexts: [] });
      return builderAgent({ entry: () => peek("reports")[0], dataset: () => peek("datasets")[0], report: () => undefined });
    },
    check: async () => {
      const { page, visuals, pages } = await pageVisuals(fixtures().reportId!, /vue d.ensemble/i);
      if (!page) return [`no page "Vue d'ensemble" (pages: ${pages.join(", ")})`];
      const types = visuals.map((v) => v.visual?.visualType);
      const count = (re: RegExp) => types.filter((t) => re.test(t ?? "")).length;
      const line = visuals.find((v) => /line|area|combo/i.test(v.visual?.visualType ?? ""));
      return [
        ...expect(count(/^card|cardVisual|multiRowCard|kpi/) >= 1 && visuals.filter((v) => /card|kpi/i.test(v.visual?.visualType)).reduce((n, v) => n + (v.visual?.query?.queryState?.Values?.projections?.length ?? v.visual?.query?.queryState?.Data?.projections?.length ?? 1), 0) >= 4, `KPI cards: ${types.join(", ")}`),
        ...expect(line && /Sales LY|LY/.test(fieldsOf(line)), "no line chart of sales vs last year"),
        ...expect(count(/donut|pie/i) >= 1, "no donut by channel"),
        ...expect(count(/slicer/i) >= 1, "no slicer"),
        ...expect(visuals.some((v) => /Product/.test(fieldsOf(v)) && (v.filterConfig || v.visual?.query?.sortDefinition)), "no top 10 products (a sorted / Top N visual on Product)"),
        ...layoutProblems(visuals, page.width ?? 1280, page.height ?? 720),
      ];
    },
  },
  {
    name: "builder-edit",
    description: "Builder (live): change the overview page (report measure, chart vs target, remove the donut)",
    request:
      "Sur la page « Vue d'ensemble » : remplace l'anneau par un histogramme des ventes par région comparées à l'objectif, et ajoute une mesure de rapport « Panier moyen HT » (ventes divisées par commandes) affichée dans une nouvelle carte.",
    setup: async () => {
      const d = await testDataset();
      const entry = await testReport(d);
      reset({ datasets: [d], reports: [entry], contexts: [] });
      return builderAgent({ entry: () => peek("reports")[0], dataset: () => peek("datasets")[0], report: () => undefined });
    },
    check: async () => {
      const id = fixtures().reportId!;
      const { page, visuals } = await pageVisuals(id, /vue d.ensemble/i);
      if (!page) return ["the page \"Vue d'ensemble\" is gone (run builder-overview first)"];
      const parts = await definition({ id });
      const ext = parts.find((p) => p.path === "definition/reportExtensions.json");
      const measures = ext ? JSON.parse(partText(ext)).entities.flatMap((e: any) => e.measures ?? []) : [];
      const basket = measures.find((m: any) => /panier/i.test(m.name));
      const regionChart = visuals.find((v) => /column|bar/i.test(v.visual?.visualType ?? "") && /Region/.test(fieldsOf(v)) && /Target/.test(fieldsOf(v)));
      return [
        ...expect(!visuals.some((v) => /donut|pie/i.test(v.visual?.visualType ?? "")), "the donut is still there"),
        ...expect(regionChart, "no column/bar chart of sales by region with the target"),
        ...expect(basket && /DIVIDE|\//.test(basket.expression), `no report measure "Panier moyen HT" (report measures: ${measures.map((m: any) => m.name).join(", ") || "none"})`),
        ...expect(basket && visuals.some((v) => /card|kpi/i.test(v.visual?.visualType ?? "") && fieldsOf(v).includes(basket.name)), "no card showing the new measure"),
        ...layoutProblems(visuals, page.width ?? 1280, page.height ?? 720),
      ];
    },
  },
  {
    name: "research-targets",
    description: "Data Scientist (live): regions behind target YTD, order date vs ship date — a complete analysis with a chart",
    request: "Fais une analyse complète : quelles régions sont en retard sur leur objectif en 2026 depuis le début de l'année ? Et quelle différence fait la date d'expédition par rapport à la date de commande sur les ventes mensuelles de 2025 ?",
    setup: async (): Promise<Agent> => {
      const d = await testDataset();
      reset({ datasets: [d], researches: [{ ...newResearch("Objectifs"), id: "r2", datasetIds: [d.id] }], reports: [], contexts: [] });
      return researchAgent(() => readResearch(peek("researches")[0]), () => peek("datasets"));
    },
    check: async () => {
      const r = readResearch(peek("researches")[0]) as ResearchEntry;
      const daxOf = (c: ResearchEntry["cards"][number]) => (c.kind === "chart" ? c.dax : JSON.stringify(c.spec));
      return [
        ...r.cards.flatMap((c) => (c.error || !c.rows?.length ? [`card "${c.title}": ${c.error ?? "no rows"}`] : [])),
        ...expect(r.cards.some((c) => /Target/i.test(daxOf(c)) && /Region/i.test(daxOf(c))), "no card of regions vs target"),
        ...expect(r.cards.some((c) => c.kind === "chart"), "no chart for the complete analysis"),
        ...expect(/r[ée]gion/i.test(r.text), "no text naming the regions"),
      ];
    },
  },
  {
    name: "dataset-question",
    description: "Datasets (live): a simple DAX question should be answered at once, without questions",
    request: "Écris un code DAX pour afficher la dernière commande.",
    setup: async () => {
      const d = await testDataset();
      reset({ datasets: [d], contexts: [] });
      return {
        instructions: `Screen: DATASETS. The user reads a semantic model: you explain it, find fields and answer with DAX.\n${DAX_RULES}`,
        tools: [daxTool(() => ({ id: d.id })), ...modelTools(() => peek("datasets")[0])],
        context: async () => withContexts([await datasetText(peek("datasets")[0])], []),
      };
    },
    check: async () => [],
  },
];
