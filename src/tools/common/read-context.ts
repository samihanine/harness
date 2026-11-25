/** A page of the chosen projects (Confluence), by its id or title: for the pages not given whole at the start. */
import { z } from "zod";
import { projectPages } from "@server/projects";
import { tool } from "../tool";

export const readProjectPage = tool({
  name: "read_project_page",
  description: "Read a Confluence page of the projects chosen in the chat (by its id or title): the pages listed at the start of the conversation without their text.",
  args: z.object({ page: z.string() }),
  async run(a, ctx) {
    const { pages } = await projectPages(ctx.scope.projects ?? []);
    const p = pages.find((x) => x.id === a.page) ?? pages.find((x) => x.title.toLowerCase() === a.page.toLowerCase()) ?? pages.find((x) => x.title.toLowerCase().includes(a.page.toLowerCase()));
    if (!p) throw new Error(`No page "${a.page}" in the chosen projects. Pages: ${pages.map((x) => `"${x.title}" (${x.id})`).join(", ") || "none"}`);
    return `# ${p.title}\n\n${p.markdown.slice(0, 60_000)}`;
  },
});
