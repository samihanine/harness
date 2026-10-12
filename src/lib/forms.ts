/** Microsoft Forms linked to an Excel table: its questions as a plain schema (read from the form link). */
import { readForm } from "@/server/ms";

export type FormQuestionType = "choice" | "text" | "long_text" | "number" | "date" | "rating" | "ranking" | "likert" | "file" | "nps" | "other";

export type FormQuestion = {
  id: string;
  title: string;
  type: FormQuestionType;
  required: boolean;
  /** choice: several answers allowed. */
  multiple?: boolean;
  options?: string[];
  allowOther?: boolean;
  /** file: accepted kinds, count and size (MB). */
  file?: { types: string[]; maxCount?: number; maxSizeMb?: number };
  /** Forms' own type and settings, as read. */
  raw: { type: string; info: Record<string, unknown> };
};

export type FormSchema = { url: string; id: string; title: string; questions: FormQuestion[]; readAt: string };

const json = (text: unknown) => {
  try {
    return typeof text === "string" && text ? (JSON.parse(text) as Record<string, any>) : {};
  } catch {
    return {};
  }
};

function question(q: Record<string, any>): FormQuestion {
  const info = json(q.questionInfo);
  const kind = String(q.type ?? "").replace(/^Question\./, "");
  const base = { id: String(q.id), title: String(q.title ?? ""), required: !!q.required, raw: { type: String(q.type), info } };
  switch (kind) {
    case "Choice":
      return {
        ...base,
        type: "choice",
        multiple: info.ChoiceType === 2,
        options: (info.Choices ?? q.choices ?? []).map((c: Record<string, unknown>) => String(c.Description ?? c.description ?? c.FormsProDisplayRTText ?? "")),
        allowOther: !!info.AllowOtherAnswer,
      };
    case "TextField":
      return { ...base, type: info.Multiline ? "long_text" : info.IsNumber || info.Restriction?.Type === 1 ? "number" : "text" };
    case "DateTime":
    case "Date":
      return { ...base, type: "date" };
    case "Rating":
      return { ...base, type: "rating", options: info.Length ? Array.from({ length: Number(info.Length) }, (_, i) => String(i + 1)) : undefined };
    case "Ranking":
      return { ...base, type: "ranking", options: (info.Choices ?? []).map((c: Record<string, unknown>) => String(c.Description ?? "")) };
    case "MatrixChoiceGroup":
    case "Likert":
      return { ...base, type: "likert" };
    case "NPS":
      return { ...base, type: "nps" };
    case "FileUpload":
      return {
        ...base,
        type: "file",
        file: { types: Object.entries(info.FileTypes ?? {}).filter(([, on]) => on).map(([t]) => t), maxCount: info.MaxFileCount, maxSizeMb: info.MaxFileSize },
      };
    default:
      return { ...base, type: "other" };
  }
}

/** Reads a form from its link (sharing or response page link); questions in their order. */
export async function readFormSchema(url: string): Promise<FormSchema> {
  const form = await readForm({ data: { url } });
  const questions = [...form.questions].sort((a, b) => Number(a.order ?? 0) - Number(b.order ?? 0)).map(question);
  return { url: url.trim(), id: form.id, title: form.title, questions, readAt: new Date().toISOString() };
}
