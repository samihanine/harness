/** Each screen's agent: its system prompt and its tools (src/tools), chosen from what the screen shows. */
import type { ChatScope, ScreenId } from "@shared/types";
import type { Tool, ToolContext } from "@/tools/tool";
import { slicerBriefing } from "@/tools/trainer/briefing";
import { readProjectPage } from "@/tools/common/read-context";
import { modelStructure } from "@/tools/model/model-structure";
import { describeTable } from "@/tools/model/describe-table";
import { searchModel } from "@/tools/model/search-model";
import { runDax } from "@/tools/model/run-dax";
import { listReports } from "@/tools/powerbi/list-reports";
import { reportStructure } from "@/tools/powerbi/report-structure";
import { lookAtReport } from "@/tools/powerbi/look-at-report";
import { readView } from "@/tools/trainer/read-view";
import { setSlicer } from "@/tools/trainer/set-slicer";
import { setFilter } from "@/tools/trainer/set-filter";
import { goToPage } from "@/tools/trainer/go-to-page";
import { resetView } from "@/tools/trainer/reset-view";
import { visualData } from "@/tools/trainer/visual-data";
import { reportOverview } from "@/tools/builder/report-overview";
import { readDefinitionFile, writeDefinitionFile } from "@/tools/builder/definition-file";
import { addPage } from "@/tools/builder/add-page";
import { addVisual } from "@/tools/builder/add-visual";
import { updateVisual } from "@/tools/builder/update-visual";
import { formatOptions, formatVisual } from "@/tools/builder/format-visual";
import { layoutPage, placeVisuals } from "@/tools/builder/layout-page";
import { removeItem } from "@/tools/builder/remove";
import { discardChanges, saveReport } from "@/tools/builder/save-report";
import { lookAtSheet, readRange, workbookInfo } from "@/tools/workbook/read";
import { addChart, conditionalFormat, createTable, dataValidation, formatRange, sheetAction, writeRange } from "@/tools/workbook/write";
import { loadDax, modelGrid } from "@/tools/workbook/model";
import { TRAINER } from "@/prompts/trainer";
import { BUILDER } from "@/prompts/builder";
import { reportMeasure } from "@/tools/builder/report-measure";
import { SCIENTIST } from "@/prompts/scientist";
import { PRESENTER } from "@/prompts/presenter";
import { teamsPrompt } from "@/prompts/teams";
import { TEAMS_TOOLS } from "@/tools/teams/teams";
import { SLIDE_TOOLS } from "@/tools/slides/slides";

/** `chatMember`: no clarification question first (an assistant in a chat answers at once, or not at all). */
export type ScreenAgent = { system: (scope: ChatScope) => string; tools: (scope: ChatScope) => Tool[]; briefing?: (ctx: ToolContext) => Promise<string | undefined>; chatMember?: boolean };

const modelTools = [modelStructure, describeTable, searchModel, runDax];
const hasModel = (s: ChatScope) => !!(s.dataset || s.report?.datasetId || s.datasets?.length);
const withContext = (s: ChatScope, tools: Tool[]) => (s.projects?.length ? [...tools, readProjectPage] : tools);

const AGENTS: Record<ScreenId, ScreenAgent> = {
  trainer: {
    system: () => TRAINER,
    briefing: slicerBriefing,
    tools: (s) => withContext(s, [readView, setSlicer, setFilter, goToPage, resetView, visualData, lookAtReport, reportStructure, listReports, ...(hasModel(s) ? modelTools : [])]),
  },
  builder: {
    system: () => BUILDER,
    tools: (s) => {
      const read = [lookAtReport, reportStructure, listReports, ...(hasModel(s) ? modelTools : [])];
      // The copy of the user's .pbix (My workspace): its definition is the draft; the model is never changed.
      if (s.report) return withContext(s, [reportOverview, addPage, addVisual, updateVisual, formatVisual, formatOptions, layoutPage, placeVisuals, removeItem, reportMeasure, saveReport, discardChanges, readDefinitionFile, writeDefinitionFile, ...read]);
      return withContext(s, read);
    },
  },
  scientist: {
    system: () => SCIENTIST,
    tools: (s) => withContext(s, [workbookInfo, readRange, lookAtSheet, writeRange, formatRange, conditionalFormat, dataValidation, sheetAction, createTable, addChart, ...(hasModel(s) ? [...modelTools, modelGrid, loadDax] : [])]),
  },
  teams: {
    chatMember: true,
    system: (s) => teamsPrompt((s.doc ?? {}) as Parameters<typeof teamsPrompt>[0]),
    // Reading only: the reports and models chosen for it, the projects' pages, the chat.
    tools: (s) => withContext(s, [...TEAMS_TOOLS, listReports, reportStructure, lookAtReport, ...(hasModel(s) ? modelTools : [])]),
  },
  presenter: {
    system: () => PRESENTER,
    tools: (s) => withContext(s, [...SLIDE_TOOLS, listReports, reportStructure, lookAtReport, ...(hasModel(s) ? modelTools : [])]),
  },
};

export const screenAgent = (screen: ScreenId) => {
  const a = AGENTS[screen];
  if (!a) throw new Error(`Unknown screen ${screen}`);
  return a;
};
