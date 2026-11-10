/**
 * Eval scenarios: seed data, the agent of a screen built as the screen builds it, a user request in
 * French (as users write), and checks on the data afterwards (empty list = passed).
 */
import type { Agent } from "@/agent/loop";
import type { ExcelTable } from "@/features/excel/sync";
import { checkValue, columnOf, DEFAULT_VIEW, type Column, type Row, type TableEntry, type View } from "@/features/excel/table";
import { excelAgent } from "@/features/excel/tools";
import { Deck } from "@/features/presenter/pptx/deck";
import { slideTypes } from "@/features/presenter/pptx/template";
import { presenterAgent } from "@/features/presenter/tools";
import { sampleDeck, templateDeck } from "../../tests/fixtures";

export type Scenario = {
  name: string;
  description: string;
  request: string;
  /** Answers to the assistant's question rounds (default: the first option of each question). */
  answers?: string[];
  setup: () => Promise<Agent>;
  check: () => Promise<string[]>;
};

const expect = (ok: unknown, problem: string) => (ok ? [] : [problem]);

/* ---------------------------------- Excel ---------------------------------- */

const TASK_COLUMNS: Column[] = [
  { name: "Title", type: "text" },
  { name: "Status", type: "option", options: [{ value: "Open", color: "blue" }, { value: "In progress", color: "amber" }, { value: "Done", color: "green" }] },
  { name: "Priority", type: "option", options: [{ value: "Haute", color: "red" }, { value: "Moyenne", color: "orange" }, { value: "Basse", color: "green" }] },
  { name: "Due", type: "date" },
];
const TASK_ROWS: Row[] = [
  { Title: "Mettre à jour le tableau de bord", Status: "Open", Priority: "Basse", Due: "2026-11-02" },
  { Title: "Relancer le fournisseur", Status: "Done", Priority: "Moyenne", Due: "2026-10-12" },
  { Title: "Préparer le comité", Status: "Open", Priority: "Moyenne", Due: "2026-11-15" },
  { Title: "Corriger les doublons clients", Status: "In progress", Priority: "Basse", Due: "2026-10-30" },
  { Title: "Former l'équipe", Status: "Open", Priority: "", Due: "" },
  { Title: "Archiver 2024", Status: "Done", Priority: "Basse", Due: "2026-09-01" },
];

/** The Excel agent on an in-memory table (the same checks of values as the screen, no Graph). */
let rows: Row[] = [];
let entry: TableEntry;
function excelSetup(): Agent {
  rows = structuredClone(TASK_ROWS);
  entry = { id: "t1", title: "Tasks", excel: { link: "x", driveId: "d", itemId: "i", name: "Tasks.xlsx", table: "Tasks", fileUrl: "x" }, columns: TASK_COLUMNS, views: [DEFAULT_VIEW], contextIds: [] };
  let view: View = DEFAULT_VIEW;
  const checked = (values: Record<string, unknown>) => Object.fromEntries(Object.entries(values).map(([k, v]) => [columnOf(entry.columns, k)!.name, checkValue(columnOf(entry.columns, k)!, v)]));
  const table = {
    get data() {
      return { header: entry.columns.map((c) => c.name), rows, cells: [] };
    },
    update: async (changes: { index: number; values: Record<string, unknown> }[]) => (changes.forEach(({ index, values }) => (rows[index] = { ...rows[index], ...checked(values) })), changes.length),
    add: async (added: Record<string, unknown>[]) => ((rows = [...rows, ...added.map(checked)]), added.map((_, i) => rows.length - added.length + i)),
    remove: async (indexes: number[]) => ((rows = rows.filter((_, i) => !indexes.includes(i))), indexes.length),
    setColumns: async (columns: Column[]) => void (entry = { ...entry, columns }),
  } as unknown as ExcelTable;
  return excelAgent(() => entry, () => table, () => view, (v) => (view = v));
}

/* -------------------------------- Presenter -------------------------------- */

let deck: Deck;
async function presenterSetup(withTemplate = true): Promise<Agent> {
  deck = await Deck.open(await sampleDeck());
  const template = await Deck.open(await templateDeck());
  const types = await slideTypes(template);
  let shown = 1;
  return presenterAgent({ deck: () => deck, template: () => (withTemplate ? { deck: template, types } : undefined), changed: async () => undefined, undo: async () => false, show: (n) => (shown = n), current: () => shown, sources: () => ({ datasets: [], reports: [] }) });
}

/* -------------------------------- Scenarios -------------------------------- */

export const SCENARIOS: Scenario[] = [
  {
    name: "excel-update",
    description: "Excel data: set the priority of every open task (by condition)",
    request: "Mets toutes les tâches ouvertes en priorité Haute.",
    setup: async () => excelSetup(),
    check: async () => {
      const wrong = rows.filter((r, i) => (r.Status === "Open" ? r.Priority !== "Haute" : r.Priority !== TASK_ROWS[i].Priority));
      return expect(!wrong.length, `wrong rows: ${wrong.map((r) => `${r.Title} (${r.Status}, ${r.Priority})`).join(", ")}`);
    },
  },
  {
    name: "excel-column",
    description: "Excel columns: add an option column with its values",
    request: "Ajoute une colonne « Équipe » avec les choix Finance, RH et IT.",
    setup: async () => excelSetup(),
    check: async () => {
      const c = columnOf(entry.columns, "Équipe");
      return [...expect(c?.type === "option", `column ${c ? `of type ${c.type}` : "missing"}`), ...expect(["Finance", "RH", "IT"].every((v) => c?.options?.some((o) => o.value === v)), `options: ${c?.options?.map((o) => o.value).join(", ")}`)];
    },
  },
  {
    name: "slides-figures",
    description: "Presenter: update the quarter and the figures of a real deck",
    request: "Passe la présentation au T3 (Q3) : le CA est de 12,4 M€ et la marge de 38 %.",
    setup: async () => presenterSetup(),
    check: async () => {
      const text = await deck.outline(true);
      return [...expect(!/Q2/.test(text), "Q2 is still in the deck"), ...expect(/12[,.]4/.test(text), "the revenue is not updated"), ...expect(/38/.test(text), "the margin is not updated"), ...expect((await deck.count()) === 3, `${await deck.count()} slides (expected 3)`)];
    },
  },
  {
    name: "slides-textbox",
    description: "Presenter: add content to a slide as a new text box (never over an existing text)",
    request: "Ajoute sur la slide 2 une zone de texte avec les types de sites : Store, Regional Warehouse, Central Warehouse.",
    setup: async () => presenterSetup(false),
    check: async () => {
      const s = await deck.slide(2);
      const texts = s.shapes.filter((x) => x.kind === "text");
      return [...expect((await deck.count()) === 3, `${await deck.count()} slides (expected 3: no duplicate)`), ...expect(texts.length === 3, `${texts.length} text shapes on slide 2 (expected 3: a new one)`), ...expect(s.title === "Results Q2", `the title became "${s.title}"`), ...expect(/Regional Warehouse/.test(texts.map((x) => x.paragraphs.map((p) => p.text).join(" ")).join(" ")), "the values are not on the slide")];
    },
  },
  {
    name: "slides-add",
    description: "Presenter: add a slide of a template type and fill it",
    request: "Ajoute avant la dernière slide une slide en deux colonnes : à gauche les points forts (croissance, marge), à droite les risques (stocks).",
    setup: async () => presenterSetup(),
    check: async () => {
      const slides = await deck.slides();
      const added = slides[slides.length - 2];
      const text = added?.shapes.map((s) => s.paragraphs.map((p) => p.text).join(" ")).join(" ") ?? "";
      return [...expect(slides.length === 4, `${slides.length} slides (expected 4)`), ...expect(/croissance/i.test(text) && /stock/i.test(text), "the new slide does not hold both columns' texts"), ...expect(!/Left column|Right column/.test(text), "the template's texts are left")];
    },
  },
];
