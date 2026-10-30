/**
 * The Confluence screen's agent: it reads and changes the open page's markdown (a draft in the screen),
 * written to Confluence when its answer is done.
 */
import { z } from "zod";
import { tool, type Agent } from "@/agent/loop";
import { CONFLUENCE_INSTRUCTIONS } from "@/prompts/confluence";

/** What the agent works on: the open page's draft, and how to save it. */
export type PageDraft = { title?: string; pages: { id: string; title: string }[]; draft: () => string; saved: () => string; setDraft: (text: string) => void; save: () => Promise<string> };

export function confluenceAgent(page: () => PageDraft): Agent {
  const open = () => {
    const p = page();
    if (!p.title) throw new Error("No page is open: ask the user to choose one in the list at the top.");
    return p;
  };
  return {
    instructions: CONFLUENCE_INSTRUCTIONS,
    context: async () => `Pages of the Confluence folder: ${page().pages.map((p) => `"${p.title}"`).join(", ") || "none"}.`,
    state: async () => {
      const p = page();
      return p.title ? `Open page: "${p.title}" (${p.draft().length} characters${p.draft() !== p.saved() ? ", unsaved changes" : ""}).` : "No page is open.";
    },
    tools: [
      tool({
        name: "read_page",
        description: "Reads the open page's markdown (with the unsaved changes).",
        args: z.object({}),
        readOnly: true,
        run: async () => open().draft() || "(empty page)",
      }),
      tool({
        name: "replace_text",
        description: "Replaces an exact passage of the page (copied from read_page) by a new text; all=true for every occurrence.",
        args: z.object({ find: z.string().min(1), replace: z.string(), all: z.boolean().optional() }),
        run: async ({ find, replace, all }) => {
          const p = open();
          const text = p.draft();
          const count = text.split(find).length - 1;
          if (!count) throw new Error("This passage is not in the page: copy it exactly from read_page (spaces and line breaks included).");
          p.setDraft(all ? text.split(find).join(replace) : text.replace(find, replace));
          return `${all ? count : 1} passage(s) replaced.`;
        },
      }),
      tool({
        name: "write_page",
        description: "Replaces the whole page by this markdown.",
        args: z.object({ markdown: z.string() }),
        run: async ({ markdown }) => (open().setDraft(markdown), "Page rewritten (draft)."),
      }),
    ],
    finish: async () => {
      const p = page();
      if (!p.title || p.draft() === p.saved()) return undefined;
      return `✓ ${await p.save()}`;
    },
  };
}
