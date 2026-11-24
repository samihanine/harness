/** The Confluence pages the user linked to the item as context (documentation, rules, definitions). */
import { z } from "zod";
import { readPageMarkdown } from "@server/confluence";
import { tool } from "../tool";

export const readContext = tool({
  name: "read_context",
  description: "Read the Confluence pages linked to this item as context (business rules, definitions, how the report is used). Give the page ids, or none for all of them.",
  args: z.object({ ids: z.array(z.string()).optional() }),
  async run(a, ctx) {
    const links = ctx.scope.contexts ?? [];
    if (!links.length) return "No context page is linked to this item.";
    const ids = a.ids?.length ? a.ids : links.map((l) => l.id);
    const pages = await Promise.all(ids.map((id) => readPageMarkdown(id).then((p) => `### ${p.title}\n${p.markdown.slice(0, 30_000)}`, (e) => `### Page ${id}\n(not readable: ${e instanceof Error ? e.message : e})`)));
    return pages.join("\n\n");
  },
});
