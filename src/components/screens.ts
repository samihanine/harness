/** The screens of the app: path, label, icon, description (home cards and the rail). */
import { BarChart3Icon, BlocksIcon, BookOpenTextIcon, DatabaseIcon, FlaskConicalIcon, GraduationCapIcon, PresentationIcon, type LucideIcon } from "lucide-react";
import type { ScreenId } from "@shared/types";

export const SCREEN_INFO: Record<ScreenId, { path: string; label: string; role: string; icon: LucideIcon; description: string }> = {
  trainer: { path: "/trainer", label: "Trainer", role: "Understand a report", icon: GraduationCapIcon, description: "Ask about a report: the AI reads it, sets the right slicers and filters, and explains what you see." },
  builder: { path: "/builder", label: "Builder", role: "Build reports", icon: BlocksIcon, description: "Create and change Power BI reports: from a sentence, a sketch or a screenshot, inspired by your reports." },
  scientist: { path: "/scientist", label: "Data Scientist", role: "Analyse data", icon: FlaskConicalIcon, description: "Combine semantic models and Excel / CSV files with SQL, and get an analysis you can export to Excel." },
  presenter: { path: "/presenter", label: "Presenter", role: "Make decks", icon: PresentationIcon, description: "Create a deck with the template from a prompt, or update an existing PowerPoint with fresh Power BI visuals." },
  reports: { path: "/reports", label: "Reports", role: "Report structures", icon: BarChart3Icon, description: "The structure of a report (pages, visuals, fields, filters), downloadable as JSON, with its context." },
  datasets: { path: "/datasets", label: "Semantic models", role: "Models and DAX", icon: DatabaseIcon, description: "Tables, relationships, measures and value examples of a model; DAX; downloadable as JSON." },
  confluence: { path: "/confluence", label: "Confluence", role: "Documentation", icon: BookOpenTextIcon, description: "Edit the documentation pages, used as context by every assistant." },
};
