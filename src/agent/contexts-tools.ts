/** The Contexts screen's agent: it creates, rewrites, edits and deletes the markdown notes given to the AI. */
import { z } from "zod";
import { tool, type Agent } from "./loop";
import { all, patch, put, remove, type ContextEntry } from "@/lib/store";

export const newContext = (name: string, text = ""): ContextEntry => ({ id: crypto.randomUUID(), name, text, updatedAt: Date.now() });
export const byName = (list: ContextEntry[], name: string) => list.find((c) => c.name.toLowerCase() === name.trim().toLowerCase());

/** The AI's tools: contexts are short markdown notes; it creates, rewrites, edits and deletes them. */
export function contextsAgent(current: () => ContextEntry | undefined, open: (id: string) => void): Agent {
  const get = async (name?: string) => {
    const list = await all("contexts");
    const c = name ? byName(list, name) : current();
    if (!c) throw new Error(name ? `No context "${name}". Contexts: ${list.map((x) => x.name).join(", ") || "none"}.` : "No context open: give its name.");
    return c;
  };
  return {
    instructions: `Screen: CONTEXTS. Contexts are markdown notes given to the AI of the other screens with the reports, datasets and Excel tables they are linked to (business rules, definitions, how to read a figure). You write and organize them: create, rewrite, edit a part, rename, delete.
- Keep them short, factual and structured (titles, lists): they are read by a small model with limited context.
- "name" omitted = the context the user has open.
- edit_context replaces an exact piece of text (copy it from read_context); rewrite_context replaces the whole text.`,
    context: async () => "",
    state: async () => {
      const list = await all("contexts");
      const c = current();
      return [`Contexts (${list.length}): ${list.map((x) => `"${x.name}" (${x.text.length} chars)`).join(", ") || "none"}.`, c ? `Open: "${c.name}".\n<open_context>\n${c.text.slice(0, 6000)}\n</open_context>` : "No context open."].join("\n\n");
    },
    tools: [
      tool({
        name: "read_context",
        description: "Reads the whole text of a context.",
        args: z.object({ name: z.string().optional() }),
        readOnly: true,
        run: async ({ name }) => {
          // Nothing open and no name: the list, so the model can pick one.
          if (!name && !current()) return `No context open. Contexts: ${(await all("contexts")).map((c) => `"${c.name}"`).join(", ") || "none"}.`;
          return (await get(name)).text || "(empty)";
        },
      }),
      tool({
        name: "create_context",
        description: "Creates a context (markdown) and opens it.",
        args: z.object({ name: z.string(), text: z.string() }),
        run: async ({ name, text }) => {
          if (byName(await all("contexts"), name)) throw new Error(`"${name}" exists: use rewrite_context.`);
          const c = await put("contexts", newContext(name.trim(), text));
          open(c.id);
          return `Context "${c.name}" created.`;
        },
      }),
      tool({
        name: "rewrite_context",
        description: "Replaces the whole text of a context (and its name when new_name is given).",
        args: z.object({ name: z.string().optional(), text: z.string(), new_name: z.string().optional() }),
        run: async ({ name, text, new_name }) => {
          const c = await get(name);
          await patch("contexts", c.id, { text, ...(new_name?.trim() && { name: new_name.trim() }), updatedAt: Date.now() });
          return `Context "${new_name ?? c.name}" saved.`;
        },
      }),
      tool({
        name: "edit_context",
        description: "Replaces an exact piece of a context's text (find) with new text (replace; empty deletes it). find empty: replace is added at the end.",
        args: z.object({ name: z.string().optional(), find: z.string(), replace: z.string() }),
        run: async ({ name, find, replace }) => {
          const c = await get(name);
          if (find && !c.text.includes(find)) throw new Error("This text is not in the context (copy it exactly from read_context).");
          const text = find ? c.text.replace(find, replace) : `${c.text.trimEnd()}\n\n${replace}`;
          await patch("contexts", c.id, { text, updatedAt: Date.now() });
          return `Context "${c.name}" edited.`;
        },
      }),
      tool({
        name: "delete_context",
        description: "Deletes a context (only when the user asks).",
        args: z.object({ name: z.string() }),
        run: async ({ name }) => {
          const c = await get(name);
          await remove("contexts", c.id);
          return `Context "${c.name}" deleted.`;
        },
      }),
    ],
  };
}

