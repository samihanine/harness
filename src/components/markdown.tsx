import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";

/**
 * The model's text made readable as markdown: the private LLM sometimes wraps it in a ```markdown
 * fence, writes "\\n" literally, or puts a list on one line (" - a - b"): fixed before rendering.
 */
export function cleanMarkdown(text: string) {
  let t = text.trim();
  const fenced = t.match(/^```(?:markdown|md)?\s*\n([\s\S]*?)\n?```$/i);
  if (fenced) t = fenced[1];
  if (!t.includes("\n") && t.includes("\\n")) t = t.replace(/\\n/g, "\n");
  // A one-line answer with list items or titles: one per line.
  if (!t.includes("\n")) t = t.replace(/\s+(?=(?:[-*•]|\d+\.)\s+\S)/g, "\n").replace(/\s+(?=#{1,4}\s)/g, "\n\n").replace(/•\s/g, "- ");
  // A list right after a sentence needs a blank line before it (or it stays in the paragraph).
  const item = (line: string) => /^\s*(?:[-*]|\d+\.)\s/.test(line);
  const lines = t.split("\n");
  return lines.map((line, i) => (i > 0 && item(line) && lines[i - 1].trim() && !item(lines[i - 1]) ? `\n${line}` : line)).join("\n");
}

export const Markdown = ({ children }: { children: string }) => (
  <div className="markdown">
    <ReactMarkdown remarkPlugins={[remarkGfm]}>{cleanMarkdown(children)}</ReactMarkdown>
  </div>
);
