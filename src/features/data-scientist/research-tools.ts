/**
 * The Data Scientist's agent: it answers business questions from the research's semantic models in the
 * canvas — a markdown text (the answer), and cards below it: a pivot table (rows, columns, values,
 * filters; the DAX is written by code) or a chart (a free DAX query), one table or chart per card.
 */
import { z } from "zod";
import { datasetText, withContexts } from "@/agent/context";
import { tool, type Agent } from "@/agent/loop";
import { DAX_RULES, daxLimits, daxTool, queryTool } from "@/agent/pbi-tools";
import { rowsText } from "@/services/pbi";
import { CHARTS, cardDax, readResearch, runCard, type Card, type ChartCard, type PivotCard, type ResearchEntry } from "@/features/data-scientist/research";
import { patch, type DatasetEntry } from "@/lib/store";
import { researchInstructions } from "@/prompts/data-scientist";

const cardId = () => crypto.randomUUID().slice(0, 8);
/** Fields: a list (a single field, or several separated by commas, are accepted too). */
const fieldList = z.preprocess((v) => (typeof v === "string" ? (v.trim() ? v.split(/,(?![^[]*\])/).map((x) => x.trim()) : []) : v), z.array(z.string()).default([]));
const filterList = z.array(z.object({ field: z.string().describe("A column 'T'[col]"), op: z.enum(["in", "not_in", ">", ">=", "<", "<=", "between", "contains"]).optional(), values: z.preprocess((v) => (v === undefined || Array.isArray(v) ? v : [v]), z.array(z.union([z.string(), z.number(), z.boolean()]))).optional(), value: z.union([z.string(), z.number()]).optional(), to: z.union([z.string(), z.number()]).optional() })).default([]);

/** The research as the AI sees it: models, the canvas text and its cards. */
function researchText(r: ResearchEntry, datasets: DatasetEntry[]) {
  const name = (id: string) => datasets.find((d) => d.id === id)?.name ?? id;
  const cardText = (c: Card) =>
    `- [${c.id}] ${c.kind === "pivot" ? `pivot "${c.title}" — rows ${c.spec.rows.join(", ") || "—"}; columns ${c.spec.columns.join(", ") || "—"}; values ${c.spec.values.join(", ")}${c.spec.filters.length ? `; ${c.spec.filters.length} filter(s)` : ""}` : `${c.chart} chart "${c.title}" — x ${c.x}, y ${c.y.join(", ")}`} on "${name(c.datasetId)}"${c.error ? ` — ERROR: ${c.error.slice(0, 120)}` : c.rows ? ` — ${c.rows.length} rows` : ""}`;
  return [
    `Models: ${r.datasetIds.map((id) => `"${name(id)}" (dataset_id ${id})`).join(", ") || "none (the user chooses them in the Datasets tab)"}.`,
    `Canvas text (${r.text.length} characters):\n${r.text ? r.text.slice(0, 1500) : "empty"}`,
    `Cards (${r.cards.length}, in order):\n${r.cards.map(cardText).join("\n") || "none"}`,
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
  const card = (id: string) => {
    const c = get().cards.find((x) => x.id === id);
    if (!c) throw new Error(`No card ${id}. Cards: ${get().cards.map((x) => x.id).join(", ") || "none"}.`);
    return c;
  };
  /** Runs the card's query, saves it (new or replacing the one of the same id) and says what it returned. */
  const put = async (c: Card, after?: string) => {
    const d = model(c.datasetId);
    const done = await runCard(c, { id: d.id, groupId: d.groupId, info: d.info });
    if (done.error) throw new Error(`${done.error}\nDAX:\n${cardDax(c, d.info)}`);
    // A replaced card keeps its place; a new one goes after `after`, else at the end.
    const old = get().cards.findIndex((x) => x.id === c.id);
    const cards = get().cards.filter((x) => x.id !== c.id);
    const anchor = after ? cards.findIndex((x) => x.id === after) : -1;
    cards.splice(old >= 0 ? old : anchor >= 0 ? anchor + 1 : cards.length, 0, done);
    await save({ cards });
    return `Card ${c.id} "${c.title}" on the canvas: ${rowsText(done.rows ?? [], 12)}`;
  };
  const pivotArgs = {
    dataset_id: z.string().optional(),
    title: z.string(),
    rows: fieldList.describe("Columns grouped as rows, e.g. ['Product'[Category]]"),
    columns: fieldList.describe("Columns grouped as columns (e.g. 'Date'[Year]), often empty"),
    values: fieldList.describe("Measures [M], or Sum('T'[col]) / Avg / Count / DistinctCount / Min / Max"),
    filters: filterList.describe("Filters on columns"),
    after: z.string().optional().describe("Card id to put it after (default: at the end)"),
  };
  return {
    autoImplement: true,
    answerOnScreen: true,
    instructions: researchInstructions(),
    context: async (focus) => withContexts(await Promise.all(models().map(async (d) => `${await datasetText(d, focus)}\n(dataset_id ${d.id})`)), models().map((d) => d.contextIds)),
    state: async () => researchText(readResearch(get()), datasets()),
    tools: [
      daxTool(() => undefined, () => models()),
      queryTool(() => undefined, () => models()),
      tool({
        name: "write_text",
        description: "Writes the canvas text (markdown, the whole text: it replaces the current one).",
        args: z.object({ text: z.string() }),
        run: async ({ text }) => (await save({ text }), "Canvas text written."),
      }),
      tool({
        name: "add_pivot",
        description: "Adds a pivot table card (or replaces card `id`): its query runs at once and its first rows are returned.",
        args: z.object({ id: z.string().optional().describe("A card to replace"), ...pivotArgs }),
        run: async ({ id, dataset_id, title, rows, columns, values, filters, after }) => {
          if (!values.length) throw new Error('Give "values" (measures).');
          const d = model(dataset_id ?? (id ? card(id).datasetId : undefined));
          const c: PivotCard = { id: id ?? cardId(), kind: "pivot", title, datasetId: d.id, spec: { rows, columns, values, filters } };
          return put(c, after);
        },
      }),
      tool({
        name: "add_chart",
        description: "Adds a chart card (or replaces card `id`) from a DAX query: x = the result column of the categories (or dates), y = the result columns of the values. Its query runs at once.",
        args: z.object({
          id: z.string().optional().describe("A card to replace"),
          dataset_id: z.string().optional(),
          title: z.string(),
          chart: z.enum(CHARTS).describe("bar (ranking), column, line / area (trend over time), pie (few parts of a whole), scatter (two measures)"),
          dax: z.string().describe("EVALUATE … returning the x column and the y columns"),
          x: z.string().describe("Result column name of the categories, e.g. Year"),
          y: fieldList.describe("Result column names of the values, e.g. ['Sales']"),
          after: z.string().optional(),
        }),
        run: async ({ id, dataset_id, title, chart, dax, x, y, after }) => {
          const d = model(dataset_id ?? (id ? card(id).datasetId : undefined));
          const strip = (n: string) => n.replace(/^.*\[(.*)\]$/, "$1");
          const c: ChartCard = { id: id ?? cardId(), kind: "chart", title, datasetId: d.id, chart, dax, x: strip(x), y: y.map(strip) };
          return put(c, after);
        },
      }),
      tool({
        name: "delete_card",
        description: "Removes a card from the canvas.",
        args: z.object({ id: z.string() }),
        run: async ({ id }) => {
          card(id);
          await save({ cards: get().cards.filter((x) => x.id !== id) });
          return `Card ${id} removed.`;
        },
      }),
      tool({
        name: "move_card",
        description: "Moves a card to a position (0 = first).",
        args: z.object({ id: z.string(), to: z.number() }),
        run: async ({ id, to }) => {
          const c = card(id);
          const cards = get().cards.filter((x) => x !== c);
          cards.splice(Math.max(0, Math.min(to, cards.length)), 0, c);
          await save({ cards });
          return `Card ${id} moved.`;
        },
      }),
    ],
  };
}
