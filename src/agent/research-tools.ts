/**
 * The Research screen's agent: answers from the research's semantic models, shown on a canvas
 * (markdown + Power BI visuals) whose visuals are the pages of one report per model (a short text and
 * ONE visual each, a matrix by default), written when the answer is done, with its own report measures.
 */
import { z } from "zod";
import { datasetText, withContexts } from "./context";
import { tool, type Agent } from "./loop";
import { DAX_RULES, daxTool, queryTool } from "./pbi-tools";
import { dax } from "@/lib/pbi";
import { CHARTS, checkPage, saveResearchReport, type Block, type MatrixPage, type ResearchEntry, type ResearchMeasure } from "@/lib/research";
import { patch, type DatasetEntry } from "@/lib/store";

const blockId = () => crypto.randomUUID().slice(0, 8);
/** Fields: a list (a single field, or several separated by commas, are accepted too). */
const fieldList = z.preprocess((v) => (typeof v === "string" ? (v.trim() ? v.split(/,(?![^[]*\])/).map((x) => x.trim()) : []) : v), z.array(z.string()).default([]));

/** The research as the AI sees it: models, canvas blocks, report pages and measures. */
function researchText(r: ResearchEntry, datasets: DatasetEntry[]) {
  const name = (id: string) => datasets.find((d) => d.id === id)?.name ?? id;
  const blockText = (b: Block) => (b.kind === "markdown" ? `- [${b.id}] text: ${b.text.replace(/\s+/g, " ").slice(0, 90)}` : b.kind === "visual" ? `- [${b.id}] visual of the page "${b.page}" (${name(b.datasetId)})` : `- [${b.id}] older DAX chart "${b.title}" (can be deleted)`);
  return [
    `Models: ${r.datasetIds.map((id) => `"${name(id)}" (dataset_id ${id})`).join(", ") || "none (the user chooses them in the Datasets tab)"}.`,
    `Canvas (${r.blocks.length} blocks, in order):\n${r.blocks.map(blockText).join("\n") || "empty"}`,
    ...r.datasetIds.map((id) => {
      const rep = r.reports[id];
      const pages = rep?.pages.length ? rep.pages.map((p) => `\n- page "${p.title}": ${p.chart ?? "matrix"}; rows ${p.rows.join(", ") || "—"}; columns ${p.columns.join(", ") || "—"}; values ${p.values.join(", ") || "—"}${p.filters.length ? `; filters ${p.filters.map((f) => `${f.field} in (${f.values.join(", ")})`).join("; ")}` : ""}${p.text ? `; text: ${p.text.replace(/\s+/g, " ").slice(0, 60)}` : ""}`).join("") : "no page";
      const measures = rep?.measures?.length ? `\n  report measures: ${rep.measures.map((m) => `'${m.table}'[${m.name}] = ${m.expression.replace(/\s+/g, " ").slice(0, 100)}`).join("; ")}` : "";
      return `Report on "${name(id)}"${rep?.id ? "" : " (not created yet)"}${rep?.dirty ? " (unsaved changes)" : ""}: ${pages}${measures}`;
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
  const insert = (blocks: Block[], b: Block, after?: string) => {
    const next = [...blocks];
    const i = after ? next.findIndex((x) => x.id === after) : -1;
    next.splice(i >= 0 ? i + 1 : next.length, 0, b);
    return next;
  };
  /** The research report's measures of a model (usable in its pages, query and run_dax). */
  const measuresOf = (datasetId: string) => get().reports[datasetId]?.measures ?? [];
  const setReport = (datasetId: string, change: { pages?: MatrixPage[]; measures?: ResearchMeasure[] }) => {
    const rep = get().reports[datasetId] ?? { pages: [] };
    return save({ reports: { ...get().reports, [datasetId]: { ...rep, ...change, dirty: true } } });
  };
  const pageArgs = z.object({
    dataset_id: z.string().optional(),
    title: z.string().describe("Page name, also the visual's title (an existing title replaces that page)"),
    chart: z.enum(CHARTS).optional().describe("The visual: matrix (default, best to read and check figures), table, bar (ranking), column, line / area (trend over time), pie (few parts of a whole), card (one or a few values)"),
    rows: fieldList.describe("Matrix rows / chart axis or slices: columns, e.g. ['Product'[Category]]"),
    columns: fieldList.describe("Matrix columns / chart series (e.g. a year), often empty"),
    values: fieldList.describe("Measures (model or report measures), or Sum('T'[col]) / Avg / Count / Min / Max"),
    filters: z.array(z.object({ field: z.string(), values: z.array(z.union([z.string(), z.number()])) })).default([]).describe("Page filters: a column and the values kept"),
    text: z.string().optional().describe("Short markdown above the visual (2–4 lines): what it shows"),
    on_canvas: z.boolean().optional().describe("Show the visual on the canvas (default true)"),
    after: z.string().optional().describe("Canvas block id to put it after (default: at the end)"),
  });
  return {
    instructions: `Screen: RESEARCH. You answer business questions from the research's semantic models and present the results:
- the CANVAS (Canvas tab): markdown blocks (findings, explanations) and Power BI visuals. Each visual is a page of the model's research REPORT (Report tab): set_page makes the page — an optional short text and ONE visual — and shows its visual on the canvas. Prefer a MATRIX (it shows the exact figures and how they split); a chart only when it says it better: line for a trend over time, bar for a ranking, pie for a few parts of a whole, card for one or a few key figures.
- This is not the Builder: one visual per page, no formatting, no dashboards. The reports are written when your answer is done; their visuals then show on the canvas.
- REPORT MEASURES (set_measure): ratios, shares, growth vs last year… that the model has not: defined in the research report (checked on the model first), usable at once in set_page values, query and run_dax.
In Research, a question on the data is answered ON THE CANVAS (visuals + a short markdown summary) — never only in the chat: it is a change, so plan it and confirm it with ask_user (one short question), with the key figures already in its message.
How to work: the context lists the models' tables, columns and measures (do not sample tables to learn them). Check the figures with query (a spec: values by groups, filtered, top N) or run_dax (raw DAX for the rest), then set_page for each visual and add_text for the summary, naming the measures used and the report page. Titles, texts and page names in the user's language. Use only fields that exist in the models.
${DAX_RULES}`,
    context: async (focus) => withContexts(await Promise.all(models().map(async (d) => `${await datasetText(d, focus)}\n(dataset_id ${d.id})`)), models().map((d) => d.contextIds)),
    state: async () => researchText(get(), datasets()),
    tools: [
      daxTool(() => undefined, () => models(), async (d) => measuresOf(d.id)),
      queryTool(() => undefined, () => models(), async (d) => measuresOf(d.id)),
      tool({
        name: "add_text",
        description: "Adds a markdown block to the canvas (after a block id, else at the end).",
        args: z.object({ text: z.string(), after: z.string().optional() }),
        run: async ({ text, after }) => {
          const b: Block = { id: blockId(), kind: "markdown", text };
          await save({ blocks: insert(get().blocks, b, after) });
          return `Text block ${b.id} added.`;
        },
      }),
      tool({
        name: "update_text",
        description: "Changes the markdown of a text block.",
        args: z.object({ id: z.string(), text: z.string() }),
        run: async ({ id, text }) => {
          const b = block(id);
          if (b.kind !== "markdown") throw new Error(`Block ${id} is not a text: change a visual with set_page (same title).`);
          await save({ blocks: get().blocks.map((x) => (x.id === id ? { ...b, text } : x)) });
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
        description: "Removes a block from the canvas (a visual's page stays in the report: delete_page for that).",
        args: z.object({ id: z.string() }),
        run: async ({ id }) => {
          block(id);
          await save({ blocks: get().blocks.filter((x) => x.id !== id) });
          return `Block ${id} removed.`;
        },
      }),
      tool({
        name: "set_page",
        description: "Adds or replaces a page of a model's report — a short text and ONE visual (rows, columns, values, page filters) — and shows its visual on the canvas.",
        args: pageArgs,
        run: async ({ dataset_id, title, chart, rows, columns, values, filters, text, on_canvas, after }) => {
          const d = model(dataset_id);
          const rep = get().reports[d.id] ?? { pages: [] };
          const old = rep.pages.find((p) => p.title === title.trim());
          // Same names as the page it replaces: the canvas keeps showing it.
          const page: MatrixPage = { title: title.trim(), chart, rows, columns, values, filters, text, name: old?.name, visualName: old?.visualName };
          checkPage(page, d.info, measuresOf(d.id));
          const pages = old ? rep.pages.map((p) => (p === old ? page : p)) : [...rep.pages, page];
          await setReport(d.id, { pages });
          const shown = get().blocks.some((b) => b.kind === "visual" && b.datasetId === d.id && b.page === page.title);
          if (on_canvas !== false && !shown) await save({ blocks: insert(get().blocks, { id: blockId(), kind: "visual", datasetId: d.id, page: page.title }, after) });
          return `Page "${page.title}" set${on_canvas !== false ? " and on the canvas" : ""} (written to the report when your answer is done).`;
        },
      }),
      tool({
        name: "delete_page",
        description: "Deletes a page of a model's report (and its visual on the canvas).",
        args: z.object({ dataset_id: z.string().optional(), title: z.string() }),
        run: async ({ dataset_id, title }) => {
          const d = model(dataset_id);
          const rep = get().reports[d.id];
          if (!rep?.pages.some((p) => p.title === title)) throw new Error(`No page "${title}".`);
          await setReport(d.id, { pages: rep.pages.filter((p) => p.title !== title) });
          await save({ blocks: get().blocks.filter((b) => !(b.kind === "visual" && b.datasetId === d.id && b.page === title)) });
          return `Page "${title}" deleted.`;
        },
      }),
      tool({
        name: "set_measure",
        description: "Creates or changes a measure of the model's research report (not of the model): checked on the model first, usable at once in set_page values, query and run_dax.",
        args: z.object({ dataset_id: z.string().optional(), table: z.string().describe("Home table (a table of the model)"), name: z.string(), expression: z.string(), format: z.string().optional().describe('e.g. "#,0", "0.0%"') }),
        run: async ({ dataset_id, table, name, expression, format }) => {
          const d = model(dataset_id);
          const others = measuresOf(d.id).filter((m) => m.name !== name);
          const next = [...others, { table, name, expression, ...(format && { format }) }];
          const defs = next.map((m) => `MEASURE '${m.table.replace(/'/g, "''")}'[${m.name}] = ${m.expression}`).join("\n");
          const [row] = await dax(d, `DEFINE\n${defs}\nEVALUATE ROW("value", [${name}])`);
          await setReport(d.id, { measures: next });
          return `Report measure '${table}'[${name}] set (value over all the data: ${row?.value ?? "blank"}).`;
        },
      }),
      tool({
        name: "delete_measure",
        description: "Deletes a measure of the research report (pages using it must change first).",
        args: z.object({ dataset_id: z.string().optional(), name: z.string() }),
        run: async ({ dataset_id, name }) => {
          const d = model(dataset_id);
          const rep = get().reports[d.id];
          if (!measuresOf(d.id).some((m) => m.name === name)) throw new Error(`No report measure "${name}".`);
          const used = rep?.pages.filter((p) => p.values.some((v) => v.includes(`[${name}]`)) || p.values.includes(name)).map((p) => `"${p.title}"`) ?? [];
          if (used.length) throw new Error(`[${name}] is used by the page(s) ${used.join(", ")}: change them first.`);
          await setReport(d.id, { measures: measuresOf(d.id).filter((m) => m.name !== name) });
          return `Report measure [${name}] deleted.`;
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

