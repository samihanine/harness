/** Raw PBIR files of the draft: list, read, write (for what the other tools do not cover). */
import { z } from "zod";
import { partText } from "@server/pbi/definitions";
import { writeJson } from "@server/pbi/draft";
import { tool } from "../tool";
import { draft } from "./shared";

export const readDefinitionFile = tool({
  name: "read_definition_file",
  description: "Read a file of the report's PBIR definition (e.g. definition/pages/<page>/visuals/<visual>/visual.json, definition/report.json, the theme). Without path: the list of files.",
  args: z.object({ path: z.string().optional() }),
  async run(a, ctx) {
    const d = await draft(ctx);
    if (!a.path) return d.parts.map((p) => p.path).join("\n");
    const p = d.parts.find((x) => x.path === a.path);
    if (!p) throw new Error(`No file ${a.path}.`);
    const text = partText(p);
    return text.length > 60_000 ? `${text.slice(0, 60_000)}\n…[cut]` : text;
  },
});

export const writeDefinitionFile = tool({
  name: "write_definition_file",
  description:
    "Write a whole PBIR file of the draft (valid JSON following its $schema): for changes the other tools do not cover (a theme, bookmarks, a visual's advanced settings, report.json settings). Read the file first and keep what you do not change.",
  args: z.object({ path: z.string().regex(/^(definition\/|StaticResources\/)/), content: z.string().describe("The file's whole JSON.") }),
  changes: true,
  async run(a, ctx) {
    const d = await draft(ctx);
    let value: unknown;
    try {
      value = JSON.parse(a.content);
    } catch (e) {
      throw new Error(`Not valid JSON: ${e instanceof Error ? e.message : e}`);
    }
    writeJson(d, a.path, value);
    return `Written ${a.path} in the draft (save_report to apply).`;
  },
});
