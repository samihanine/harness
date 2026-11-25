/** Each screen's agent: its system prompt and its tools (src/tools), chosen from what the screen shows. */
import type { ChatScope, ScreenId } from "@shared/types";
import type { Tool, ToolContext } from "@/tools/tool";
import { slicerBriefing } from "@/tools/trainer/briefing";
import { readProjectPage } from "@/tools/common/read-context";
import { EDITOR_TOOLS } from "@/tools/editor/editor";
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
import { layoutPage } from "@/tools/builder/layout-page";
import { removeItem } from "@/tools/builder/remove";
import { discardChanges, saveReport } from "@/tools/builder/save-report";
import { deleteMeasure, saveMeasure } from "@/tools/datasets/measures";
import { listSources } from "@/tools/scientist/sources";
import { loadDax, loadFile } from "@/tools/scientist/load";
import { createStep, sqlQuery } from "@/tools/scientist/sql";
import { readAnalysis, removeBlock, setBlock, setTitle } from "@/tools/scientist/blocks";
import { capturePowerbi } from "@/tools/deck/capture-powerbi";
import { editSlides, readDeck, writeDeck } from "@/tools/deck/deck";
import { lookAtDeck } from "@/tools/deck/look-at-deck";
import { lookAtSlides, pptxStructure, viewPicture } from "@/tools/pptx/structure";
import { addPicture, addTemplateSlide, addText, replacePicture, setText, shapeAction, slideAction } from "@/tools/pptx/edit";
import { TRAINER } from "@/prompts/trainer";
import { BUILDER, BUILDER_EDITOR } from "@/prompts/builder";
import { SCIENTIST } from "@/prompts/scientist";
import { PRESENTER_CREATE, PRESENTER_EDIT } from "@/prompts/presenter";

export type ScreenAgent = { system: (scope: ChatScope) => string; tools: (scope: ChatScope) => Tool[]; briefing?: (ctx: ToolContext) => Promise<string | undefined> };

const modelTools = [modelStructure, describeTable, searchModel, runDax];
const hasModel = (s: ChatScope) => !!(s.dataset || s.report?.datasetId || s.datasets?.length);
const withContext = (s: ChatScope, tools: Tool[]) => (s.projects?.length ? [...tools, readProjectPage] : tools);

const AGENTS: Record<ScreenId, ScreenAgent> = {
  trainer: {
    system: () => TRAINER,
    briefing: slicerBriefing,
    tools: (s) => withContext(s, [readView, setSlicer, setFilter, goToPage, resetView, visualData, lookAtReport, reportStructure, ...(hasModel(s) ? modelTools : [])]),
  },
  builder: {
    system: (s) => (s.report?.access === "editor" ? BUILDER_EDITOR : BUILDER),
    tools: (s) => {
      const read = [lookAtReport, reportStructure, listReports, ...(hasModel(s) ? [...modelTools, saveMeasure] : [])];
      // Editable definition: the PBIR tools; editable in the Power BI editor only: the editor tools; else reading.
      if (s.report?.access === "edit") return withContext(s, [reportOverview, addPage, addVisual, updateVisual, formatVisual, formatOptions, layoutPage, removeItem, saveReport, discardChanges, readDefinitionFile, writeDefinitionFile, ...read]);
      if (s.report?.access === "editor") return withContext(s, [...EDITOR_TOOLS, ...read]);
      return withContext(s, read);
    },
  },
  scientist: {
    system: () => SCIENTIST,
    tools: (s) => withContext(s, [listSources, loadDax, loadFile, sqlQuery, createStep, setBlock, removeBlock, readAnalysis, setTitle, ...(hasModel(s) ? [...modelTools, saveMeasure, deleteMeasure] : [])]),
  },
  presenter: {
    system: (s) => (s.mode === "edit" ? PRESENTER_EDIT : PRESENTER_CREATE),
    tools: (s) =>
      withContext(s, [
        ...(s.mode === "edit" ? [pptxStructure, lookAtSlides, viewPicture, setText, replacePicture, addPicture, addText, shapeAction, slideAction, addTemplateSlide] : [readDeck, writeDeck, editSlides, lookAtDeck]),
        capturePowerbi,
        listReports,
        reportStructure,
        lookAtReport,
        ...(hasModel(s) ? modelTools : []),
      ]),
  },
};

export const screenAgent = (screen: ScreenId) => {
  const a = AGENTS[screen];
  if (!a) throw new Error(`Unknown screen ${screen}`);
  return a;
};
