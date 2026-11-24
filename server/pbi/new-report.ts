/** A new, empty report (one page) live-connected to a semantic model, written in PBIR: the Builder fills it. */
import type { Ref } from "@shared/types";
import { fabric, workspaceId } from "../ms/api";
import { textPart } from "./definitions";
import { buildPage, newName } from "./visual-spec";

const json = (path: string, value: unknown) => textPart(path, JSON.stringify(value, null, 2));

export async function createBlankReport(name: string, dataset: Ref, groupId?: string) {
  const ws = await workspaceId(groupId ?? dataset.groupId);
  const page = newName();
  const parts = [
    json("definition.pbir", { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json", version: "4.0", datasetReference: { byConnection: { connectionString: `semanticmodelid=${dataset.id}` } } }),
    json("definition/version.json", { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json", version: "2.0.0" }),
    json("definition/report.json", { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json", themeCollection: { baseTheme: { name: "CY24SU10", reportVersionAtImport: "5.61", type: "SharedResources" } }, layoutOptimization: "None", settings: { useStylableVisualContainerHeader: true, defaultDrillFilterOtherVisuals: true } }),
    json("definition/pages/pages.json", { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json", pageOrder: [page], activePageName: page }),
    json(`definition/pages/${page}/page.json`, buildPage(page, "Page 1")),
  ];
  const created = await fabric(`/workspaces/${ws}/reports`, "POST", { displayName: name, definition: { parts } });
  const id = (created?.id as string | undefined) ?? (await fabric(`/workspaces/${ws}/reports`)).value.find((r: { displayName: string }) => r.displayName === name)?.id;
  if (!id) throw new Error("The report was created but could not be found: look for it in the workspace.");
  return { id: id as string, groupId: groupId ?? dataset.groupId };
}
