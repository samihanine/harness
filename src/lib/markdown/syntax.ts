/**
 * The markdown of the app (pages, guides, answers): GitHub markdown plus a few constructs that
 * Confluence can hold, each written to a Confluence macro and read back from it (storage.ts):
 * - code fences with a language → code macro (colored on screen: highlight.tsx);
 * - ```mermaid → a diagram (components/diagram.tsx), attached to the page as a picture, its source kept in an expand macro;
 * - `> [!NOTE] title` / TIP / WARNING / IMPORTANT → colored panels (info / tip / warning / note macros), EXPAND → expand macro;
 * - `` `status:green:Done` `` → a status lozenge;
 * - `![](icon:database)` → an inline icon (a lucide icon, attached as a small picture);
 * - `- [ ] task` → a Confluence task list.
 */

export const PANEL_KINDS = { NOTE: "info", TIP: "tip", WARNING: "warning", IMPORTANT: "note", EXPAND: "expand" } as const;
export type PanelKind = keyof typeof PANEL_KINDS;
export const STATUS_COLORS = ["green", "blue", "red", "yellow", "grey", "purple"] as const;

/** `> [!NOTE] title` blockquotes → ```panel-note fences (title on the first line): one block for both renderers. */
export function withPanels(md: string) {
  return md.replace(/^> \[!(NOTE|TIP|WARNING|IMPORTANT|EXPAND)\][ \t]*(.*)\n((?:>.*(?:\n|$))*)/gm, (_, kind: string, title: string, body: string) => {
    const inner = body.replace(/^> ?/gm, "").trimEnd();
    return `\`\`\`\`panel-${kind.toLowerCase()}\n${title.trim()}\n${inner}\n\`\`\`\`\n`;
  });
}

/** A panel fence's kind, title and markdown. */
export function panelOf(lang: string, code: string) {
  const kind = lang.match(/^panel-(\w+)$/)?.[1]?.toUpperCase() as PanelKind | undefined;
  if (!kind || !(kind in PANEL_KINDS)) return undefined;
  const [title, ...rest] = code.split("\n");
  return { kind, title: title.trim(), body: rest.join("\n") };
}

/** `status:green:Done` (inline code) → its color and text. */
export const statusOf = (code: string) => {
  const m = code.match(/^status:(\w+):(.+)$/);
  return m ? { color: m[1].toLowerCase(), text: m[2] } : undefined;
};

/** A short stable hash of a text (attachment names: the same diagram, the same file). */
export function hashOf(text: string) {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) h = Math.imul(h ^ text.charCodeAt(i), 0x01000193);
  return (h >>> 0).toString(16).padStart(8, "0");
}

export const diagramFile = (code: string) => `hub-diagram-${hashOf(code.trim())}.png`;
export const iconFile = (name: string) => `hub-icon-${name}.png`;

/** The rules the AI gets to write a page (screens that write markdown). */
export const MARKDOWN_RULES = `Markdown (GitHub) with these additions, all kept in Confluence:
- Diagrams: a \`\`\`mermaid flowchart (flowchart LR or TD; subgraph … end; a --> b, a -->|label| b, -.-> dashed; colors with :::blue / :::green / :::amber / :::violet / :::teal / :::red / :::orange / :::pink / :::gray; an icon before a label: id[icon:database Sales]; a detail after <br/>).
- Code: fences with a language (ts, dax, json, sql, bash, python, m).
- Panels: a quote starting with > [!NOTE] title (or TIP, WARNING, IMPORTANT, EXPAND for a foldable block), its lines starting with >.
- Status: \`status:green:Done\` (green, blue, red, yellow, grey, purple).
- Icons in text: ![](icon:database) (lucide icon names).
- Task lists: - [ ] to do / - [x] done.
No HTML.`;
