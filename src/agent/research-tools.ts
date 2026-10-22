/**
 * The Research screen's agent: answers from the research's semantic models, shown as a canvas
 * (markdown + charts computed by DAX) and as one matrix report per model (written when it is done).
 */
import { z } from "zod";
import { datasetText, withContexts } from "./context";
import { tool, type Agent } from "./loop";
import { DAX_RULES, daxTool } from "./pbi-tools";
import { rowsText } from "@/lib/pbi";
import { CHARTS, checkPage, computeBlock, saveResearchReport, type Block, type MatrixPage, type ResearchEntry } from "@/lib/research";
import { patch, type DatasetEntry } from "@/lib/store";

type Chart = Extract<Block, { kind: "chart" }>;
const blockId = () => crypto.randomUUID().slice(0, 8);
/** Fields: a list (a single field, or several separated by commas, are accepted too). */
const fieldList = z.preprocess((v) => (typeof v === "string" ? (v.trim() ? v.split(/,(?![^[]*\])/).map((x) => x.trim()) : []) : v), z.array(z.string()).default([]));

/** The research as the AI sees it: models, canvas blocks, report pages. */
function researchText(r: ResearchEntry, datasets: DatasetEntry[]) {
  const name = (id: string) => datasets.find((d) => d.id === id)?.name ?? id;
  return [
    `Models: ${r.datasetIds.map((id) => `"${name(id)}" (dataset_id ${id})`).join(", ") || "none (the user chooses them in the Datasets tab)"}.`,
    `Canvas (${r.blocks.length} blocks, in order):\n${r.blocks.map((b) => (b.kind === "markdown" ? `- [${b.id}] text: ${b.text.replace(/\s+/g, " ").slice(0, 90)}` : `- [${b.id}] ${b.chart} chart "${b.title}" on ${name(b.datasetId)}${b.error ? ` — ERROR: ${b.error.slice(0, 160)}` : ` — ${b.rows?.length ?? 0} rows`}`)).join("\n") || "empty"}`,
    ...r.datasetIds.map((id) => {
      const rep = r.reports[id];
      return `Report on "${name(id)}"${rep?.id ? "" : " (not created yet)"}${rep?.dirty ? " (unsaved changes)" : ""}: ${rep?.pages.length ? rep.pages.map((p) => `\n- page "${p.title}": rows ${p.rows.join(", ") || "—"}; columns ${p.columns.join(", ") || "—"}; values ${p.values.join(", ") || "—"}${p.filters.length ? `; filters ${p.filters.map((f) => `${f.field} in (${f.values.join(", ")})`).join("; ")}` : ""}`).join("") : "no page"}`;
    }),
  ].join("\n\n");
}

export function researchAgent(get: () => ResearchEntry, datasets: () => DatasetEntry[]): Agent {
  const save = (change: Partial<ResearchEntry>) => patch("researches", get().id, { ...change, updatedAt: Date.now() });
  const models = () => datasets().filter((d) => get().datasetIds.includes(d.id));
  const model = (id?: string) => {
    const list = models();
    const d = id ? list.find((x) => x.id === id || x.name.toLowerCase() === id.toLowerCase()) : list.length === 1 ? list[0] : undefined;
    if (!d) throw new Error(list.length ? `Give dataset_id: one of ${list.map((x) => `${x.id} (${x.name})`).join(", ")}.` : "No model in this research: ask the user to choose one in the Datasets tab.");
    return d;
  };
  const block = (id: string) => {
    const b = get().blocks.find((x) => x.id === id);
    if (!b) throw new Error(`No block ${id}. Blocks: ${get().blocks.map((x) => x.id).join(", ") || "none"}.`);
    return b;
  };
  const insert = (b: Block, after?: string) => {
    const blocks = [...get().blocks];
    const i = after ? blocks.findIndex((x) => x.id === after) : -1;
    blocks.splice(i >= 0 ? i + 1 : blocks.length, 0, b);
    return blocks;
  };
  /** A chart is kept only when its query works (the error goes back to the AI to fix). */
  const computed = async (b: Chart) => {
    const d = model(b.datasetId);
    const done = await computeBlock({ ...b, datasetId: d.id }, d);
    if (done.error) throw new Error(`The DAX query failed: ${done.error}`);
    if (!done.rows?.length) throw new Error("The query returned no row: check the filters.");
    return done;
  };
  const pageArgs = z.object({
    dataset_id: z.string().optional(),
    title: z.string().describe("Page name (an existing title replaces that page)"),
    rows: fieldList.describe("Columns as matrix rows, e.g. ['Product'[Category]]"),
    columns: fieldList.describe("Columns as matrix columns (e.g. a year), often empty"),
    values: fieldList.describe("Measures, or Sum('T'[col]) / Avg / Count / Min / Max"),
    filters: z.array(z.object({ field: z.string(), values: z.array(z.union([z.string(), z.number()])) })).default([]).describe("Page filters: a column and the values kept"),
  });
  return {
    instructions: `Screen: RESEARCH. You answer business questions from the research's semantic models and present the results in two places:
- the CANVAS (Canvas tab): text blocks (markdown: findings, explanations) and charts, each computed by one DAX query. Chart kinds: ${CHARTS.join(", ")}. The query's first column is the category (x axis / slices), the next columns are the numeric series; kpi shows the first row's numbers; table shows all rows. Keep queries small (TOPN, SUMMARIZECOLUMNS, ≤ 30 rows for charts).
- the REPORTS (Report tab): one Power BI report per model; each page is ONE full-page matrix: rows, columns, values, page filters — no formatting, no other visual. set_page adds or replaces a page; the reports are written when your answer is done.
In Research, a question on the data is answered ON THE CANVAS (charts + a short summary) and in the report — never only in the chat: it is a change, so plan it and confirm it with ask_user (one short question), with the key figures already in its message.
ALWAYS, when your answer uses data: add at least one report page (set_page) whose matrix reproduces the figures you show — the same measures (values), columns (rows / columns) and filters — so the user can see how they were obtained and check them; in the canvas summary, name the measures used (e.g. [Total Sales], [Margin %]) and the report page. The DAX of each chart is kept with it.
How to work: the context lists the models' tables, columns and measures (do not sample tables to learn them). Check the figures with run_dax, then add the charts and a short markdown summary of what they show. add_chart runs its query and returns the rows: do not run the same query again first. Titles, texts and page names in the user's language. Use only fields that exist in the models.
${DAX_RULES}`,
    context: async () => withContexts(await Promise.all(models().map(async (d) => `${await datasetText(d)}\n(dataset_id ${d.id})`)), models().map((d) => d.contextIds)),
    state: async () => researchText(get(), datasets()),
    tools: [
      daxTool(() => undefined, () => models()),
      tool({
        name: "add_text",
        description: "Adds a markdown block to the canvas (after a block id, else at the end).",
        args: z.object({ text: z.string(), after: z.string().optional() }),
        run: async ({ text, after }) => {
          const b: Block = { id: blockId(), kind: "markdown", text };
          await save({ blocks: insert(b, after) });
          return `Text block ${b.id} added.`;
        },
      }),
      tool({
        name: "add_chart",
        description: "Adds a chart computed by a DAX query (EVALUATE …) to the canvas. Returns the first rows.",
        args: z.object({ title: z.string(), chart: z.enum(CHARTS), dax: z.string(), dataset_id: z.string().optional(), after: z.string().optional() }),
        run: async ({ title, chart, dax, dataset_id, after }) => {
          const b = await computed({ id: blockId(), kind: "chart", title, chart, dax, datasetId: model(dataset_id).id });
          await save({ blocks: insert(b, after) });
          return `Chart ${b.id} added (${b.rows!.length} rows):\n${rowsText(b.rows!, 8)}`;
        },
      }),
      tool({
        name: "update_block",
        description: "Changes a block: text (markdown block), or title / chart / dax (chart, recomputed).",
        args: z.object({ id: z.string(), text: z.string().optional(), title: z.string().optional(), chart: z.enum(CHARTS).optional(), dax: z.string().optional() }),
        run: async ({ id, text, title, chart, dax }) => {
          const b = block(id);
          const next: Block = b.kind === "markdown" ? { ...b, text: text ?? b.text } : dax && dax !== b.dax ? await computed({ ...b, title: title ?? b.title, chart: chart ?? b.chart, dax }) : { ...b, title: title ?? b.title, chart: chart ?? b.chart };
          await save({ blocks: get().blocks.map((x) => (x.id === id ? next : x)) });
          return `Block ${id} updated.`;
        },
      }),
      tool({
        name: "move_block",
        description: "Moves a block to a position (0 = first).",
        args: z.object({ id: z.string(), to: z.number() }),
        run: async ({ id, to }) => {
          const b = block(id);
          const blocks = get().blocks.filter((x) => x !== b);
          blocks.splice(Math.max(0, Math.min(to, blocks.length)), 0, b);
          await save({ blocks });
          return `Block ${id} moved.`;
        },
      }),
      tool({
        name: "delete_block",
        description: "Deletes a canvas block.",
        args: z.object({ id: z.string() }),
        run: async ({ id }) => {
          block(id);
          await save({ blocks: get().blocks.filter((x) => x.id !== id) });
          return `Block ${id} deleted.`;
        },
      }),
      tool({
        name: "set_page",
        description: "Adds or replaces a page of a model's report: one full-page matrix (rows, columns, values, page filters).",
        args: pageArgs,
        run: async ({ dataset_id, title, rows, columns, values, filters }) => {
          const d = model(dataset_id);
          const page: MatrixPage = { title: title.trim(), rows, columns, values, filters };
          checkPage(page, d.info);
          const rep = get().reports[d.id] ?? { pages: [] };
          const pages = rep.pages.some((p) => p.title === page.title) ? rep.pages.map((p) => (p.title === page.title ? page : p)) : [...rep.pages, page];
          await save({ reports: { ...get().reports, [d.id]: { ...rep, pages, dirty: true } } });
          return `Page "${page.title}" set (written to the report when your answer is done).`;
        },
      }),
      tool({
        name: "delete_page",
        description: "Deletes a page of a model's report.",
        args: z.object({ dataset_id: z.string().optional(), title: z.string() }),
        run: async ({ dataset_id, title }) => {
          const d = model(dataset_id);
          const rep = get().reports[d.id];
          if (!rep?.pages.some((p) => p.title === title)) throw new Error(`No page "${title}".`);
          await save({ reports: { ...get().reports, [d.id]: { ...rep, pages: rep.pages.filter((p) => p.title !== title), dirty: true } } });
          return `Page "${title}" deleted.`;
        },
      }),
    ],
    finish: async () => {
      const done = await publishReports(get(), datasets());
      return done.length ? `✓ Report${done.length > 1 ? "s" : ""} written: ${done.join(", ")}.` : undefined;
    },
  };
}

/** Writes the reports whose pages changed (created the first time). */
export async function publishReports(r: ResearchEntry, datasets: DatasetEntry[]) {
  const done: string[] = [];
  for (const id of r.datasetIds) {
    const rep = r.reports[id];
    const d = datasets.find((x) => x.id === id);
    if (!rep?.dirty || !d || !rep.pages.length) continue;
    const saved = await saveResearchReport(`${r.name} — ${d.name}`, d, rep, rep.pages);
    await patch("researches", r.id, { reports: { ...r.reports, [id]: saved } });
    r = { ...r, reports: { ...r.reports, [id]: saved } };
    done.push(d.name);
  }
  return done;
}

