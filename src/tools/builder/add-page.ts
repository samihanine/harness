/** A new page in the report being built. */
import { z } from "zod";
import { readJson, writeJson } from "@server/pbi/draft";
import { buildPage, newName } from "@server/pbi/visual-spec";
import { tool } from "../tool";
import { draft } from "./shared";

export const addPage = tool({
  name: "add_page",
  description: "Add a page (default 1280×720) to the report being built. Returns its name (id) for the other tools.",
  args: z.object({ display_name: z.string(), width: z.number().optional(), height: z.number().optional(), after: z.string().optional().describe("Place it after this page (default: at the end).") }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    const name = newName();
    writeJson(d, `definition/pages/${name}/page.json`, buildPage(name, a.display_name, a.width ?? 1280, a.height ?? 720));
    const pagesPath = "definition/pages/pages.json";
    const pages = d.parts.some((p) => p.path === pagesPath) ? readJson(d, pagesPath) : { $schema: "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json", pageOrder: [] };
    const order: string[] = pages.pageOrder ?? [];
    const at = a.after ? order.indexOf(a.after) : -1;
    at >= 0 ? order.splice(at + 1, 0, name) : order.push(name);
    writeJson(d, pagesPath, { ...pages, pageOrder: order, activePageName: pages.activePageName ?? name });
    return `Page "${a.display_name}" added: name=${name}.`;
  },
});
