/**
 * Eval scenarios: seed data, the agent of a screen built as the screen builds it, a user request in
 * French (as users write), and checks on the data afterwards (empty list = passed).
 */
import type { Agent } from "@/agent/loop";
import { contextsAgent } from "@/features/contexts/contexts-tools";
import { excelData } from "@/features/excel/excel-tools";
import { slidesAgent } from "@/features/presenter/slides-tools";
import type { Table } from "@/services/excel";
import type { Field } from "@/features/excel/schema";
import type { ContextEntry, DeckEntry, SchemaEntry } from "@/lib/store";
import { EMPTY_VIEW } from "@/features/excel/table-view";
import type { Write } from "@/features/excel/use-table";
import { EXAMPLE_DECK, checkSlide } from "@/features/presenter/slides";
import * as store from "@/lib/store";

/** The in-memory store loaded in place of the real one (preload.ts). */
const { all, peek, reset } = store as unknown as typeof import("./fake-store");

export type Scenario = {
  name: string;
  description: string;
  request: string;
  /** Answers to the planner's question rounds (default: the first option of each question). */
  answers?: string[];
  setup: () => Promise<Agent>;
  check: () => Promise<string[]>;
};

const expect = (ok: unknown, problem: string) => (ok ? [] : [problem]);
const deck = async () => ((await all("decks")) as DeckEntry[])[0];

/** The slides agent reading the in-memory deck at call time, as the screen does. */
async function slidesSetup(entry: DeckEntry) {
  reset({ decks: [entry], reports: [] });
  let shown = 0;
  return slidesAgent(() => peek("decks")[0], (i) => (shown = i), () => shown);
}

/* ---------------------------------- Excel ---------------------------------- */

const TASK_FIELDS: Field[] = [
  { name: "id", type: "string", label: "ID" },
  { name: "Title", type: "string", required: true },
  { name: "Status", type: "option", options: [{ value: "Open", color: "blue" }, { value: "In progress", color: "amber" }, { value: "Done", color: "green" }] },
  { name: "Priority", type: "option", options: [{ value: "Haute", color: "red" }, { value: "Moyenne", color: "orange" }, { value: "Basse", color: "green" }] },
  { name: "Due", type: "date" },
];
const TASK_ROWS = [
  { id: "TASK_001", Title: "Mettre à jour le tableau de bord", Status: "Open", Priority: "Basse", Due: "2026-11-02" },
  { id: "TASK_002", Title: "Relancer le fournisseur", Status: "Done", Priority: "Moyenne", Due: "2026-10-12" },
  { id: "TASK_003", Title: "Préparer le comité", Status: "Open", Priority: "Moyenne", Due: "2026-11-15" },
  { id: "TASK_004", Title: "Corriger les doublons clients", Status: "In progress", Priority: "Basse", Due: "2026-10-30" },
  { id: "TASK_005", Title: "Former l'équipe", Status: "Open", Priority: null, Due: null },
  { id: "TASK_006", Title: "Archiver 2024", Status: "Done", Priority: "Basse", Due: "2026-09-01" },
];
const TASKS: SchemaEntry = { id: "s1", title: "Tasks", excel: { link: "x", driveId: "d", itemId: "i", name: "Tasks.xlsx", table: "Tasks", fileUrl: "x" }, fields: TASK_FIELDS, contextIds: [] };

/** The data agent on an in-memory table: writes apply their optimistic change only (no Graph). */
let table: Table;
function excelSetup(): Agent {
  reset({ schemas: [TASKS], contexts: [] });
  table = { fields: TASK_FIELDS, rows: structuredClone(TASK_ROWS), header: TASK_FIELDS.map((f) => f.name) };
  let view = EMPTY_VIEW;
  const write = (async (_task, { optimistic } = {}) => void (optimistic && (table = optimistic(table)))) as Write;
  // Made again at each call, as the screen does at each render (the table changes).
  const fresh = () => excelData({ schema: TASKS, table, shown: () => table.rows.map((_, i) => i), checked: [], view, setView: (v) => (view = { ...view, ...v }), openRow: () => undefined, write });
  const first = fresh();
  return {
    instructions: first.instructions,
    context: () => fresh().context(),
    state: () => fresh().state!(),
    tools: first.tools.map((t) => ({ ...t, run: (args: unknown) => fresh().tools.find((x) => x.name === t.name)!.run(args) })),
  };
}

/* -------------------------------- Scenarios -------------------------------- */

export const SCENARIOS: Scenario[] = [
  {
    name: "contexts-new",
    description: "Contexts: write a context of sales KPI definitions",
    request: "Écris un contexte « KPI ventes » avec les définitions du chiffre d'affaires, de la marge et du panier moyen.",
    setup: async () => {
      reset({ contexts: [] });
      return contextsAgent(() => undefined, () => undefined);
    },
    check: async () => {
      const c = ((await all("contexts")) as ContextEntry[]).find((x) => /kpi/i.test(x.name));
      return [...expect(c, "no context named “KPI…”"), ...expect(c && /marge/i.test(c.text) && /panier/i.test(c.text), "the context lacks the margin or average basket definition")];
    },
  },
  {
    name: "slides-new",
    description: "Slides: write a 5-6 slide deck from figures",
    request: "Fais une présentation de 6 slides sur les résultats du trimestre : CA 12,4 M€ (+8 %), marge 38 %, 1 240 commandes. Termine par les prochaines étapes.",
    setup: async () => {
      return slidesSetup({ id: "d1", name: "Résultats", deck: { ...EXAMPLE_DECK, title: "Résultats", slides: [] }, updatedAt: 0 });
    },
    check: async () => {
      const slides = (await deck()).deck.slides;
      const invalid = slides.flatMap((s, i) => {
        try {
          checkSlide(s);
          return [];
        } catch (e) {
          return [`slide ${i + 1}: ${e instanceof Error ? e.message : e}`];
        }
      });
      return [
        ...expect(slides.length >= 5 && slides.length <= 8, `${slides.length} slides (asked 6)`),
        ...expect(["hero", "title", "cover"].includes(String(slides[0]?.template)), "does not start with a cover slide"),
        ...expect(["end", "title", "cover"].includes(String(slides.at(-1)?.template)) || /étape|next/i.test(JSON.stringify(slides.slice(-2))), "does not end with the next steps / a closing slide"),
        ...expect(/12[,.]4/.test(JSON.stringify(slides)), "the revenue figure is missing"),
        ...invalid,
      ];
    },
  },
  {
    name: "slides-edit",
    description: "Slides: change the theme color and delete a slide",
    request: "Passe la couleur principale en vert et supprime la slide « Priorities ».",
    setup: async () => {
      return slidesSetup({ id: "d1", name: "Review", deck: structuredClone(EXAMPLE_DECK), updatedAt: 0 });
    },
    check: async () => {
      const d = (await deck()).deck;
      const [r, g, b] = (d.theme.primary.replace("#", "").match(/../g) ?? []).map((x) => parseInt(x, 16));
      return [...expect(g > r && g > b, `primary color ${d.theme.primary} is not green`), ...expect(!d.slides.some((s) => s.title === "Priorities"), "the Priorities slide is still there"), ...expect(d.slides.length === EXAMPLE_DECK.slides.length - 1, `${d.slides.length} slides (expected ${EXAMPLE_DECK.slides.length - 1})`)];
    },
  },
  {
    name: "excel-update",
    description: "Excel data: set the priority of every open task",
    request: "Mets toutes les tâches ouvertes en priorité Haute.",
    setup: async () => excelSetup(),
    check: async () => {
      const wrong = table.rows.filter((r, i) => (r.Status === "Open" ? r.Priority !== "Haute" : r.Priority !== TASK_ROWS[i].Priority));
      return expect(!wrong.length, `wrong rows: ${wrong.map((r) => `${r.id} (${r.Status}, ${r.Priority})`).join(", ")}`);
    },
  },

];
