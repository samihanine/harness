/** The Confluence page being edited (a draft in the screen, published by the user). */
import { z } from "zod";
import { listPages } from "@server/confluence";
import { docOf, setDoc } from "../helpers";
import { tool } from "../tool";

type Draft = { pageId: string; title: string; markdown: string; version: number };

export const readPage = tool({
  name: "read_page",
  description: "The page being edited, as markdown (the draft on screen, unpublished changes included).",
  args: z.object({}),
  async run(_, ctx) {
    const d = docOf<Draft>(ctx);
    if (!d) throw new Error("No page is open.");
    return `# ${d.title}\n\n${d.markdown}`;
  },
});

export const writePage = tool({
  name: "write_page",
  description: "Replace the draft of the page with new markdown (headings, lists, tables, code, links). The user reviews and publishes it.",
  args: z.object({ markdown: z.string(), title: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const d = docOf<Draft>(ctx);
    if (!d) throw new Error("No page is open.");
    setDoc(ctx, { ...d, markdown: a.markdown, title: a.title ?? d.title });
    return "Draft updated: the user sees it and publishes it with Publish.";
  },
});

export const listConfluencePages = tool({
  name: "list_pages",
  description: "The pages of the Confluence folder (title, id).",
  args: z.object({}),
  async run() {
    return (await listPages()).map((p) => `- "${p.title}" (${p.id})`).join("\n") || "No page.";
  },
});
