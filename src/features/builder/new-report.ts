/**
 * A new Power BI report written by code (PBIR, through the Fabric definition API), live-connected to a
 * semantic model: the Builder starts from it ("New report from a semantic model").
 */
import { all } from "@/lib/store";
import type { Ref } from "@/services/pbi";
import { fabric } from "@/services/ms";
import { definition, textPayload, type Part } from "@/services/report-def";
import { workspaceId } from "@/services/tmdl";
import { buildPage, newName } from "./visual-spec";

const REPORT_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json";
const VERSION_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json";
const PAGES_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json";

const json = (path: string, value: unknown): Part => ({ path, payload: textPayload(JSON.stringify(value, null, 2)), payloadType: "InlineBase64" });


/**
 * The connection to the model: copied from a report of the app on the same model when there is one
 * (exactly what Power BI wrote), else the short form Fabric accepts.
 */
async function connection(dataset: Ref): Promise<Part> {
  const report = (await all("reports")).find((r) => r.datasetId === dataset.id && r.access === "edit");
  if (report) {
    const part = await definition(report)
      .then((parts) => parts.find((p) => p.path === "definition.pbir"))
      .catch(() => undefined);
    if (part) return part;
  }
  // The $schema is required: without it Fabric reads the old format (and asks for five more properties).
  return json("definition.pbir", { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json", version: "4.0", datasetReference: { byConnection: { connectionString: `semanticmodelid=${dataset.id}` } } });
}

/**
 * A new report in PBIR format (editable by the Builder's tools), live-connected to a semantic model,
 * in a workspace (default: My workspace); returns its id.
 */
async function createReport(name: string, dataset: Ref, pagesParts: Part[], groupId?: string) {
  const ws = await workspaceId(groupId);
  const parts = [
    await connection(dataset),
    json("definition/version.json", { $schema: VERSION_SCHEMA, version: "2.0.0" }),
    json("definition/report.json", { $schema: REPORT_SCHEMA, themeCollection: { baseTheme: { name: "CY24SU10", reportVersionAtImport: "5.61", type: "SharedResources" } }, layoutOptimization: "None", settings: { useStylableVisualContainerHeader: true, defaultDrillFilterOtherVisuals: true } }),
    ...pagesParts,
  ];
  const created = await fabric(`/workspaces/${ws}/reports`, "POST", { displayName: name, definition: { parts } });
  const id = (created?.id as string | undefined) ?? (await fabric(`/workspaces/${ws}/reports`)).value.find((r: { displayName: string }) => r.displayName === name)?.id;
  if (!id) throw new Error("The report was created but could not be found: look for it in the workspace.");
  return id as string;
}

/** An empty report (one blank page) on a model, in My workspace: the Builder's AI fills it. */
export function createBlankReport(name: string, dataset: Ref) {
  const page = newName();
  return createReport(name, dataset, [
    json("definition/pages/pages.json", { $schema: PAGES_SCHEMA, pageOrder: [page], activePageName: page }),
    json(`definition/pages/${page}/page.json`, buildPage(page, "Page 1")),
  ]);
}

