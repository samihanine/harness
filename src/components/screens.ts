/** The screens of the app: path, label, description (home cards and the rail); their icons: screen-icons.tsx. */
import type { ScreenId } from "@shared/types";

export const SCREEN_INFO: Record<ScreenId, { path: string; label: string; role: string; description: string }> = {
  trainer: { path: "/trainer", label: "Trainer", role: "Understand a report", description: "Ask about a report: the AI reads it, sets the right slicers and filters, and explains what you see." },
  builder: { path: "/builder", label: "Builder", role: "Build reports", description: "Drop a .pbix: the AI changes it with every option of Power BI Desktop, shown as Power BI draws it; download it connected." },
  scientist: { path: "/scientist", label: "Data Scientist", role: "Analyse data", description: "An Excel workbook built with you: connected to the Power BI models, with your files, formulas and formats." },
  presenter: { path: "/presenter", label: "Presenter", role: "Make decks", description: "Presentations from your templates: content in markdown, Power BI views live, PowerPoint and PDF." },
  teams: { path: "/teams", label: "Teams", role: "Answer in Teams", description: "An assistant in one of your Teams chats: it reads every message and answers when it concerns its projects and data." },
};
