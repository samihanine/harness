/**
 * Markdown ⇄ Confluence storage format. Pages are edited as markdown (by the user or the AI): the
 * page's XHTML is turned into markdown when it is read, and the markdown into XHTML (react-markdown,
 * the app's own renderer) when it is written. Confluence macros other than code blocks are kept as
 * their text only.
 */
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";

/** Markdown → storage XHTML (well-formed: void tags closed). */
export const toStorage = (markdown: string) => renderToStaticMarkup(createElement(ReactMarkdown, { remarkPlugins: [remarkGfm] }, markdown));

/** Storage XHTML → markdown: headings, paragraphs, emphasis, links, lists, code, tables, rules. */
export function toMarkdown(html: string) {
  // CDATA (the text of code macros) is not HTML: turned into escaped text first.
  const escaped = html.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, (_, text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"));
  const doc = new DOMParser().parseFromString(`<body>${escaped}</body>`, "text/html");
  return block(doc.body).replace(/\n{3,}/g, "\n\n").trim();
}

const tag = (n: Element) => n.tagName.toLowerCase();

/** Inline content: text with **bold**, *italic*, `code` and [links](url). */
function inline(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) return (node.textContent ?? "").replace(/\s+/g, " ");
  if (!(node instanceof Element)) return "";
  const inner = () => [...node.childNodes].map(inline).join("");
  switch (tag(node)) {
    case "strong":
    case "b":
      return `**${inner().trim()}**`;
    case "em":
    case "i":
      return `*${inner().trim()}*`;
    case "code":
      return `\`${node.textContent ?? ""}\``;
    case "a":
      return `[${inner().trim()}](${node.getAttribute("href") ?? ""})`;
    case "br":
      return "  \n";
    case "img":
      return `![${node.getAttribute("alt") ?? ""}](${node.getAttribute("src") ?? ""})`;
    default:
      return inner();
  }
}

/** Block content, one block per paragraph. */
function block(node: Element, depth = 0): string {
  return [...node.childNodes]
    .map((child) => {
      if (!(child instanceof Element)) return inline(child).trim() ? `${inline(child).trim()}\n\n` : "";
      const t = tag(child);
      if (/^h[1-6]$/.test(t)) return `${"#".repeat(Number(t[1]))} ${inline(child).trim()}\n\n`;
      if (t === "p") return `${inline(child).trim()}\n\n`;
      if (t === "hr") return "---\n\n";
      if (t === "blockquote") return `${block(child).trim().split("\n").map((l) => `> ${l}`).join("\n")}\n\n`;
      if (t === "pre") return `\`\`\`\n${(child.textContent ?? "").replace(/\n+$/, "")}\n\`\`\`\n\n`;
      if (t === "ul" || t === "ol") return `${list(child, depth)}\n`;
      if (t === "table") return `${table(child)}\n\n`;
      // A code macro: its text as a fenced block; another macro: its text.
      if (t === "ac:structured-macro") {
        const body = child.getElementsByTagName("ac:plain-text-body")[0]?.textContent;
        return body !== undefined ? `\`\`\`\n${body.replace(/\n+$/, "")}\n\`\`\`\n\n` : block(child, depth);
      }
      return block(child, depth);
    })
    .join("");
}

function list(node: Element, depth: number): string {
  const ordered = tag(node) === "ol";
  return [...node.children]
    .filter((li) => tag(li) === "li")
    .map((li, i) => {
      const nested = [...li.children].filter((c) => ["ul", "ol"].includes(tag(c)));
      const text = [...li.childNodes].filter((c) => !(c instanceof Element && ["ul", "ol"].includes(tag(c)))).map((c) => (c instanceof Element && tag(c) === "p" ? inline(c) : inline(c))).join("").trim();
      return `${"  ".repeat(depth)}${ordered ? `${i + 1}.` : "-"} ${text}\n${nested.map((n) => list(n, depth + 1)).join("")}`;
    })
    .join("");
}

function table(node: Element): string {
  const rows = [...node.querySelectorAll("tr")].map((tr) => [...tr.children].map((c) => inline(c).trim().replace(/\|/g, "\\|")));
  if (!rows.length) return "";
  const width = Math.max(...rows.map((r) => r.length));
  const line = (r: string[]) => `| ${Array.from({ length: width }, (_, i) => r[i] ?? "").join(" | ")} |`;
  return [line(rows[0]), `| ${Array.from({ length: width }, () => "---").join(" | ")} |`, ...rows.slice(1).map(line)].join("\n");
}
