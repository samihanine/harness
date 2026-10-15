/**
 * Microsoft Forms linked to a schema. Forms has no public API: the questions are read and written
 * through the endpoints the Forms designer uses (`formsApi` in src/server/ms.ts) — the user allowed
 * the schema editor to change their forms this way. Fragile: a change on Microsoft's side may break it.
 */
import { formsApi, readForm } from "@/server/ms";
import { COLORS, isIdField, type Field, type FieldType, type FormSection } from "./schema";
import type { FormRef } from "./store";

const SECTION = "Question.ColumnGroup";
type Question = { id: string; title: string; type: string; order: number; required: boolean; questionInfo: string; subtitle?: string | null };

const json = (text: unknown) => {
  try {
    return typeof text === "string" && text ? (JSON.parse(text) as Record<string, any>) : {};
  } catch {
    return {};
  }
};

/** Reads a form from its link: its id and title. */
export async function readFormRef(url: string): Promise<FormRef & { questions: Record<string, any>[] }> {
  const form = await readForm({ data: { url } });
  // The designer's list has the sections too (owners only); the respondents' one has not.
  const design = await formsApi({ data: { form: form.id, path: "forms('{id}')/questions" } }).then((r) => r.value as Record<string, any>[], () => undefined);
  return { url: url.trim(), id: form.id, title: form.title, questions: [...(design ?? form.questions)].sort((a, b) => Number(a.order ?? 0) - Number(b.order ?? 0)) };
}

/** A question as a field type (likert, ranking, NPS… have no field of their own: short text). */
function fieldOf(q: Record<string, any>): Omit<Field, "name"> {
  const info = json(q.questionInfo);
  const kind = String(q.type ?? "").replace(/^Question\./, "");
  const base = { label: String(q.title ?? "") || undefined, required: !!q.required || undefined, description: q.subtitle || undefined };
  const type = (t: FieldType, extra: Partial<Field> = {}) => ({ ...base, type: t, ...extra });
  switch (kind) {
    case "Choice": {
      const options = (info.Choices ?? q.choices ?? []).map((c: Record<string, unknown>) => String(c.Description ?? c.description ?? "")).filter(Boolean);
      return type("option", { multiple: info.ChoiceType === 2 || undefined, options: options.map((value: string, i: number) => ({ value, color: COLORS[(i + 1) % COLORS.length] })) });
    }
    case "TextField":
      return type(info.Multiline ? "text" : info.Restriction?.Type === "Number" || info.IsNumber ? "number" : "string");
    case "DateTime":
    case "Date":
      return type("date");
    case "Rating":
      return type("rating", { max: Number(info.Length) === 10 ? 10 : 5 });
    case "FileUpload": {
      // Images only: an image field.
      const kinds = Object.entries(info.FileTypes ?? {}).filter(([, on]) => on).map(([k]) => k);
      return type(info.HasSpecificFileType && kinds.length === 1 && kinds[0] === "Image" ? "image" : "file");
    }
    default:
      return type("string");
  }
}

/** A field name from a question title (only the characters a field name allows). */
const nameOf = (title: string) => {
  const name = title.normalize("NFC").replace(/[^\p{L}\p{N}_ ]+/gu, " ").replace(/\s+/g, " ").trim().slice(0, 60);
  return /^[\p{L}_]/u.test(name) ? name : `q ${name}`.trim();
};

/**
 * Field types and options from a form, for the columns of an Excel table: a column takes the type of
 * the question with its name (the Excel file of a form has one column per question title). Columns
 * without question keep their own field.
 */
export function withFormTypes(fields: Field[], questions: Record<string, any>[]) {
  const asked = questions.filter((q) => q.type !== SECTION);
  return fields.map((f) => {
    const q = asked.find((x) => nameOf(String(x.title ?? "")).toLowerCase() === f.name.toLowerCase() || String(x.title ?? "").toLowerCase() === f.name.toLowerCase());
    return q && !isIdField(f) ? { ...f, ...fieldOf(q), name: f.name, label: f.label } : f;
  });
}

/**
 * Where a branch goes in Forms: a section's question id, or the end of the form.
 * ASSUMPTION (to confirm on a form made in the Forms designer): the end is written "EndOfForm".
 */
export const END_OF_FORM = "EndOfForm";

/** The question a field becomes (type and settings in Forms' own format). Its title IS the column name: Forms names the Excel columns after the questions. */
type QuestionDef = { title: string; type: string; required: boolean; questionInfo: string; subtitle: string | null };
function questionOf(f: Field, target: (section: string) => string): QuestionDef {
  const branch = (option: string) => {
    const to = f.branches?.[option];
    return to ? { BranchInfo: { TargetQuestionId: to === "end" ? END_OF_FORM : target(to) } } : {};
  };
  const info: Record<FieldType, [string, Record<string, unknown>]> = {
    string: ["Question.TextField", { Multiline: false }],
    text: ["Question.TextField", { Multiline: true }],
    number: ["Question.TextField", { Multiline: false, Restriction: { Type: "Number", RestrictionType: "None" } }],
    date: ["Question.DateTime", { Date: true }],
    option: ["Question.Choice", { Choices: (f.options ?? []).map((o) => ({ Description: o.value, ...branch(o.value) })), ChoiceType: f.multiple ? 2 : 1, AllowOtherAnswer: false, OptionDisplayStyle: "ListAll" }],
    rating: ["Question.Rating", { Length: f.max ?? 5, RatingShape: "Star", MinRatingLabel: "", MaxRatingLabel: "" }],
    url: ["Question.TextField", { Multiline: false }],
    image: ["Question.FileUpload", { HasSpecificFileType: true, FileTypes: { Word: false, Excel: false, PowerPoint: false, PDF: false, Image: true, Video: false, Audio: false }, MaxFileCount: 1, MaxFileSize: 10 }],
    file: ["Question.FileUpload", { HasSpecificFileType: false, FileTypes: { Word: true, Excel: true, PowerPoint: true, PDF: true, Image: true, Video: true, Audio: true }, MaxFileCount: 1, MaxFileSize: 10 }],
  };
  const [type, questionInfo] = info[f.type];
  return { title: f.name, type, required: !!f.required, questionInfo: JSON.stringify(questionInfo), subtitle: f.description ?? null };
}

/** A section's question (Forms calls sections "ColumnGroup"), with where it goes when done. */
const sectionOf = (x: FormSection, target: (section: string) => string): QuestionDef => ({
  title: x.title,
  type: "Question.ColumnGroup",
  required: false,
  subtitle: x.description ?? null,
  questionInfo: JSON.stringify(x.next ? { BranchInfo: { TargetQuestionId: x.next === "end" ? END_OF_FORM : target(x.next) } } : {}),
});

/** What changes in a question: compared on its meaning (Forms adds its own keys to questionInfo). */
function differs(q: Question, next: QuestionDef) {
  if (q.title !== next.title || q.type !== next.type || !!q.required !== next.required || (q.subtitle ?? null) !== next.subtitle) return true;
  const a = json(q.questionInfo), b = json(next.questionInfo);
  return Object.keys(b).some((k) => JSON.stringify(a[k]) !== JSON.stringify(b[k])) || (!!a.BranchInfo !== !!b.BranchInfo);
}

/** One item of the form, in order: a section or a question, with the id it has (or will have). */
type Item = { kind: "section"; section: FormSection; id: string; existing?: Question } | { kind: "field"; field: Field; id: string; existing?: Question };
export type FormPlan = { items: Item[]; create: string[]; update: string[]; remove: Question[]; fields: Field[]; sections: FormSection[] };

const api = (form: FormRef, path: string, method?: string, body?: unknown) => formsApi({ data: { form: form.id, path, method, body } });
const newId = () => `r${crypto.randomUUID().replace(/-/g, "")}`;

/**
 * What writing the schema to its form would do: the questions (title = column name) and the sections
 * in order, with the branches; questions to create, to change, and to delete (matching no field or
 * section). A field keeps its question id (`question`); else it is matched by title. Fields with
 * `form: false` and the id column are never asked.
 */
export async function planForm(form: FormRef, fields: Field[], sections: FormSection[] = []): Promise<FormPlan> {
  const questions: Question[] = (await api(form, "forms('{id}')/questions")).value;
  const used = new Set<string>();
  const match = (id: string | undefined, title: string, type?: string) => {
    const q = questions.find((x) => !used.has(x.id) && x.id === id) ?? questions.find((x) => !used.has(x.id) && x.title.toLowerCase() === title.toLowerCase() && (type ? x.type === type : x.type !== "Question.ColumnGroup"));
    if (q) used.add(q.id);
    return q;
  };
  const asked = fields.filter((f) => !isIdField(f) && f.form !== false);
  const placed = new Set(sections.flatMap((x) => x.fields));
  const items: Item[] = [];
  const addField = (f: Field) => {
    const existing = match(f.question, f.name);
    items.push({ kind: "field", field: f, id: existing?.id ?? newId(), existing });
  };
  asked.filter((f) => !placed.has(f.name)).forEach(addField);
  for (const x of sections) {
    const existing = match(x.question, x.title, "Question.ColumnGroup");
    items.push({ kind: "section", section: x, id: existing?.id ?? newId(), existing });
    x.fields.map((n) => asked.find((f) => f.name === n)).filter(Boolean).forEach((f) => addField(f!));
  }
  const target = (title: string) => (items.find((i) => i.kind === "section" && i.section.title === title) as Item | undefined)?.id ?? END_OF_FORM;
  const plan: FormPlan = { items, create: [], update: [], remove: questions.filter((q) => !used.has(q.id)), fields: [], sections: [] };
  for (const i of items) {
    const next = i.kind === "field" ? questionOf(i.field, target) : sectionOf(i.section, target);
    const label = i.kind === "field" ? i.field.name : `Section "${i.section.title}"`;
    if (!i.existing) plan.create.push(label);
    else if (differs(i.existing, next)) plan.update.push(label);
  }
  plan.fields = fields.map((f) => {
    const i = items.find((x) => x.kind === "field" && x.field === f);
    return i ? { ...f, question: i.id } : f;
  });
  plan.sections = sections.map((x) => ({ ...x, question: items.find((i) => i.kind === "section" && i.section === x)!.id }));
  return plan;
}

/** Writes a plan to the form, in order (orders like the designer's: 1000500, 2000500…). */
export async function applyForm(form: FormRef, plan: FormPlan, progress?: (label: string, done: number, total: number) => void) {
  const total = plan.remove.length + plan.items.length;
  let done = 0;
  // Sections are "descriptive questions": changed and deleted through their own collection.
  const path = (type: string, id: string) => `forms('{id}')/${type === "Question.ColumnGroup" ? "descriptiveQuestions" : "questions"}('${id}')`;
  for (const q of plan.remove) {
    progress?.(`Deleting "${q.title}"`, done++, total);
    await api(form, path(q.type, q.id), "DELETE");
  }
  const target = (title: string) => (plan.items.find((i) => i.kind === "section" && i.section.title === title) as Item | undefined)?.id ?? END_OF_FORM;
  let order = 1000500;
  for (const i of plan.items) {
    const next = { ...(i.kind === "field" ? questionOf(i.field, target) : sectionOf(i.section, target)), order };
    order += 1000000;
    progress?.(`${i.existing ? "Checking" : "Adding"} "${next.title}"`, done++, total);
    if (!i.existing) await api(form, i.kind === "section" ? "forms('{id}')/descriptiveQuestions" : "forms('{id}')/questions", "POST", { id: i.id, ...next });
    else if (differs(i.existing, next) || i.existing.order !== order - 1000000) await api(form, path(next.type, i.id), "PATCH", next);
  }
  return { fields: plan.fields, sections: plan.sections };
}

/** A form's title as the schema's (the form keeps its other settings). */
export const renameForm = (form: FormRef, title: string) => api(form, "forms('{id}')", "PATCH", { title });

/** Link to edit a form in the Forms designer. */
export const designUrl = (form: FormRef) => `https://forms.cloud.microsoft/Pages/DesignPageV2.aspx?subpage=design&id=${form.id}`;
export const responseUrl = (form: FormRef) => `https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=${form.id}`;

/** Fields of a new schema read from a form alone (one per question), e.g. to start from a form. */
export const fieldsFromQuestions = (questions: Record<string, any>[]): Field[] =>
  questions.filter((q) => q.type !== SECTION).map((q) => ({ name: nameOf(String(q.title ?? "Question")), ...fieldOf(q), question: String(q.id) }));

/**
 * The sections and branches of a form (questions in order: a section takes the questions after it),
 * for the fields matched by name: where each section and each choice option goes.
 */
export function layoutFromQuestions(questions: Record<string, any>[], fields: Field[]) {
  const ordered = [...questions].sort((a, b) => Number(a.order ?? 0) - Number(b.order ?? 0));
  const titleOf = (id?: string) => (id === END_OF_FORM ? "end" : ordered.find((q) => q.id === id && q.type === SECTION)?.title);
  const fieldOfQuestion = (q: Record<string, any>) => fields.find((f) => f.question === q.id || f.name.toLowerCase() === nameOf(String(q.title ?? "")).toLowerCase());
  const sections: FormSection[] = [];
  const next = fields.map((f) => ({ ...f }));
  for (const q of ordered) {
    if (q.type === SECTION) {
      const to = titleOf(json(q.questionInfo).BranchInfo?.TargetQuestionId);
      sections.push({ title: String(q.title ?? "Section"), description: q.subtitle || undefined, fields: [], question: String(q.id), ...(to && { next: to }) });
      continue;
    }
    const f = fieldOfQuestion(q);
    if (!f) continue;
    if (sections.length) sections[sections.length - 1].fields.push(f.name);
    const branches = Object.fromEntries(
      (json(q.questionInfo).Choices ?? [])
        .map((c: Record<string, any>) => [String(c.Description ?? ""), titleOf(c.BranchInfo?.TargetQuestionId)])
        .filter(([, to]: [string, string | undefined]) => to),
    );
    const target = next.find((x) => x.name === f.name)!;
    if (Object.keys(branches).length) target.branches = branches;
    target.question = String(q.id);
  }
  return { fields: next, sections };
}
