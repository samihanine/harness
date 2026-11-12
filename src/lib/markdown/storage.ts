/**
 * Markdown ⇄ Confluence storage format, without loss for what the app writes: every construct of
 * syntax.ts has a macro whose parameters keep it, so a page written and read again gives the same
 * markdown. Written with the app's own markdown renderer (react-markdown): the special blocks become
 * placeholders while rendering, replaced by their macro afterwards. Read with the browser's DOM parser.
 * The files to attach (diagrams, icons) are listed: the Confluence screen draws and uploads them.
 */
import { createElement, isValidElement, type ReactElement, type ReactNode } from "react";
import ReactMarkdown, { type Components } from "react-markdown";
import remarkGfm from "remark-gfm";
import { renderXhtml } from "./render";
import { PANEL_KINDS, diagramFile, iconFile, panelOf, safeUrl, statusOf, withPanels, type PanelKind } from "./syntax";

export type StorageFile = { name: string; kind: "diagram"; code: string } | { name: string; kind: "icon"; icon: string };

/** Fence aliases → the code macro's language names, and back. */
const LANGUAGES: Record<string, string> = { ts: "typescript", tsx: "typescript", js: "javascript", jsx: "javascript", sh: "bash", shell: "bash", py: "python", yml: "yaml" };
const FENCES: Record<string, string> = { typescript: "ts", javascript: "js" };
/** Languages Confluence does not know: a close one colors them, the macro's title names them (and brings them back). */
const CUSTOM: Record<string, { language: string; title: string }> = { dax: { language: "sql", title: "DAX" }, m: { language: "text", title: "Power Query (M)" }, tmdl: { language: "yaml", title: "TMDL" } };

const cdata = (s: string) => `<![CDATA[${s.replace(/]]>/g, "]]]]><![CDATA[>")}]]>`;
const escape = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const macro = (name: string, params: Record<string, string | undefined>, body = "") =>
  `<ac:structured-macro ac:name="${name}">${Object.entries(params)
    .filter(([, v]) => v)
    .map(([k, v]) => `<ac:parameter ac:name="${k}">${escape(v!)}</ac:parameter>`)
    .join("")}${body}</ac:structured-macro>`;
const codeMacro = (code: string, language?: string, title?: string) => macro("code", { title, language }, `<ac:plain-text-body>${cdata(code)}</ac:plain-text-body>`);

/** The code element of a <pre> (its language and text). */
function fenceOf(children: ReactNode) {
  const code = (Array.isArray(children) ? children : [children]).find(isValidElement) as ReactElement<{ className?: string; children?: ReactNode }> | undefined;
  return { lang: code?.props.className?.match(/language-([\w-]+)/)?.[1] ?? "", code: String(code?.props.children ?? "").replace(/\n$/, "") };
}

/** Markdown → storage XHTML and the files it refers to. */
export function toStorage(md: string): { storage: string; files: StorageFile[] } {
  // What replaces a placeholder is made after the rendering (a panel's content is converted then, never inside a render).
  const held: (string | (() => string))[] = [];
  const files: StorageFile[] = [];
  const hold = (xml: string | (() => string)) => createElement("hub-x", { "data-i": held.push(xml) - 1 });
  const file = (f: StorageFile) => void (files.some((x) => x.name === f.name) || files.push(f));
  const inner = (text: string) => {
    const r = toStorage(text);
    r.files.forEach(file);
    return r.storage;
  };
  const components: Components = {
    pre: ({ children }) => {
      const { lang, code } = fenceOf(children);
      if (lang === "mermaid") {
        const name = diagramFile(code);
        file({ name, kind: "diagram", code });
        return hold(`<p><ac:image ac:align="center"><ri:attachment ri:filename="${name}" /></ac:image></p>${macro("expand", { title: "Diagram source" }, `<ac:rich-text-body>${codeMacro(code, "text", "hub:mermaid")}</ac:rich-text-body>`)}`);
      }
      const panel = panelOf(lang, code);
      if (panel) return hold(() => macro(PANEL_KINDS[panel.kind], { title: panel.title || undefined }, `<ac:rich-text-body>${inner(panel.body)}</ac:rich-text-body>`));
      const custom = CUSTOM[lang];
      return hold(custom ? codeMacro(code, custom.language, custom.title) : codeMacro(code, LANGUAGES[lang] ?? (lang || undefined)));
    },
    code: ({ children, className }) => {
      const status = !className && statusOf(String(children));
      if (status) return hold(macro("status", { colour: status.color[0].toUpperCase() + status.color.slice(1), title: status.text }));
      return createElement("code", null, children);
    },
    img: ({ src, alt }) => {
      const icon = String(src ?? "").match(/^icon:([\w-]+)$/)?.[1];
      if (icon) {
        file({ name: iconFile(icon), kind: "icon", icon });
        return hold(`<ac:image ac:height="16" ac:alt="${escape(alt || icon)}"><ri:attachment ri:filename="${iconFile(icon)}" /></ac:image>`);
      }
      const attachment = String(src ?? "").match(/^attachment:(.+)$/)?.[1];
      return hold(`<ac:image ac:alt="${escape(alt ?? "")}">${attachment ? `<ri:attachment ri:filename="${escape(attachment)}" />` : `<ri:url ri:value="${escape(String(src ?? ""))}" />`}</ac:image>`);
    },
    // Task lists: Confluence's own (a checkbox each), written as marks here and turned into macros below.
    ul: ({ children, className }) => createElement(className?.includes("contains-task-list") ? "hub-tasks" : "ul", null, children),
    li: ({ children, className, node }) => {
      if (!className?.includes("task-list-item")) return createElement("li", null, children);
      // Its checkbox (maybe inside a paragraph of a loose list), the box itself not written (input: below).
      return createElement("hub-task", { "data-done": JSON.stringify(node ?? {}).includes('"checked":true') ? "1" : "0" }, children);
    },
    input: () => null,
  };
  let html = renderXhtml(createElement(ReactMarkdown, { remarkPlugins: [remarkGfm], components, urlTransform: (url: string, key: string) => safeUrl(url, key === "src") }, withPanels(md)));
  // Tasks: innermost first (a task may hold a list of tasks).
  for (let i = 0; i < 10 && /<hub-task[ >]/.test(html); i++)
    html = html.replace(/<hub-task data-done="([01])">((?:(?!<hub-task[ >])[\s\S])*?)<\/hub-task>/g, (_, done: string, body: string) => `<ac:task><ac:task-status>${done === "1" ? "complete" : "incomplete"}</ac:task-status><ac:task-body>${body.replace(/^<p>([\s\S]*)<\/p>$/, "$1").trim()}</ac:task-body></ac:task>`);
  html = html.replace(/<hub-tasks>/g, "<ac:task-list>").replace(/<\/hub-tasks>/g, "</ac:task-list>");
  // Placeholders may hold others (a panel's content): replaced until none is left.
  for (let i = 0; i < 20 && html.includes("<hub-x"); i++) html = html.replace(/<hub-x data-i="(\d+)"\s*(?:\/>|><\/hub-x>)/g, (_, n: string) => (typeof held[Number(n)] === "function" ? (held[Number(n)] as () => string)() : (held[Number(n)] as string)));
  return { storage: html, files };
}

/* ----------------------------------- Read ----------------------------------- */

/** Storage XHTML → markdown: the app's constructs from their macros, the rest as close as markdown says it. */
export function fromStorage(html: string) {
  // CDATA (the text of code macros) is not HTML: turned into escaped text first.
  const escaped = html.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, (_, text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"));
  const doc = new DOMParser().parseFromString(`<body>${escaped}</body>`, "text/html");
  return block(doc.body).replace(/\n{3,}/g, "\n\n").trim();
}

const tag = (n: Element) => n.tagName.toLowerCase();
const param = (m: Element, name: string) => [...m.children].find((c) => tag(c) === "ac:parameter" && c.getAttribute("ac:name") === name)?.textContent?.trim() ?? "";
const bodyOf = (m: Element, name: string) => [...m.children].find((c) => tag(c) === name);
const PANEL_OF: Record<string, PanelKind> = { info: "NOTE", tip: "TIP", warning: "WARNING", note: "IMPORTANT", expand: "EXPAND", panel: "NOTE" };

/** An image: an icon of the app, a diagram's picture (dropped: its source follows), an attachment or a link. */
function image(node: Element) {
  const attachment = node.getElementsByTagName("ri:attachment")[0]?.getAttribute("ri:filename");
  const url = node.getElementsByTagName("ri:url")[0]?.getAttribute("ri:value");
  const alt = node.getAttribute("ac:alt") ?? "";
  if (attachment?.startsWith("hub-diagram-")) return "";
  const icon = attachment?.match(/^hub-icon-([\w-]+)\.png$/)?.[1];
  if (icon) return `![${alt === icon ? "" : alt}](icon:${icon})`;
  return attachment ? `![${alt}](attachment:${attachment})` : url ? `![${alt}](${url})` : "";
}

/** Inline content: text with **bold**, *italic*, ~~struck~~, `code`, [links](url), icons, statuses. */
function inline(node: Node): string {
  if (node.nodeType === Node.TEXT_NODE) return (node.textContent ?? "").replace(/\s+/g, " ");
  if (!(node instanceof Element)) return "";
  const inner = () => [...node.childNodes].map(inline).join("");
  switch (tag(node)) {
    case "strong":
    case "b":
      return inner().trim() ? `**${inner().trim()}**` : "";
    case "em":
    case "i":
      return inner().trim() ? `*${inner().trim()}*` : "";
    case "s":
    case "del":
      return `~~${inner().trim()}~~`;
    case "code":
      return `\`${node.textContent ?? ""}\``;
    case "a":
      return `[${inner().trim()}](${node.getAttribute("href") ?? ""})`;
    case "br":
      return "  \n";
    case "img":
      return `![${node.getAttribute("alt") ?? ""}](${node.getAttribute("src") ?? ""})`;
    case "ac:image":
      return image(node);
    case "ac:emoticon":
      // Self-closing in XML, not for the HTML parser: what follows it ends up inside.
      return (node.getAttribute("ac:emoji-fallback") ?? "") + inner();
    case "ac:link":
      return node.getElementsByTagName("ac:plain-text-link-body")[0]?.textContent ?? node.getElementsByTagName("ri:page")[0]?.getAttribute("ri:content-title") ?? inner();
    case "time":
      return node.getAttribute("datetime") ?? inner();
    case "ac:structured-macro":
      if (node.getAttribute("ac:name") === "status") return `\`status:${(param(node, "colour") || "grey").toLowerCase()}:${param(node, "title")}\``;
      return inner();
    case "ac:parameter":
      return "";
    default:
      return inner();
  }
}

/** A macro as a block of markdown (code, diagram, panel), undefined for the others. */
function macroBlock(node: Element): string | undefined {
  const name = node.getAttribute("ac:name") ?? "";
  if (name === "code" || name === "noformat") {
    const text = (bodyOf(node, "ac:plain-text-body")?.textContent ?? "").replace(/\n+$/, "");
    const title = param(node, "title");
    const custom = Object.entries(CUSTOM).find(([, c]) => c.title === title)?.[0];
    const language = param(node, "language");
    const lang = title === "hub:mermaid" ? "mermaid" : (custom ?? FENCES[language] ?? (language === "none" ? "" : language));
    return `\`\`\`${lang}\n${text}\n\`\`\`\n\n`;
  }
  const body = bodyOf(node, "ac:rich-text-body");
  // A diagram's source, in its expand macro.
  if (name === "expand" && body) {
    const code = [...body.children].find((c) => tag(c) === "ac:structured-macro" && c.getAttribute("ac:name") === "code" && param(c, "title") === "hub:mermaid");
    if (code && body.children.length === 1) return macroBlock(code);
  }
  if (PANEL_OF[name] && body) {
    const text = block(body).trim();
    return `> [!${PANEL_OF[name]}]${param(node, "title") ? ` ${param(node, "title")}` : ""}\n${text.split("\n").map((l) => (l ? `> ${l}` : ">")).join("\n")}\n\n`;
  }
  return undefined;
}

/** Block content, one block per paragraph. */
function block(node: Element, depth = 0): string {
  return [...node.childNodes]
    .map((child) => {
      if (!(child instanceof Element)) return inline(child).trim() ? `${inline(child).trim()}\n\n` : "";
      const t = tag(child);
      if (/^h[1-6]$/.test(t)) return `${"#".repeat(Number(t[1]))} ${inline(child).trim()}\n\n`;
      if (t === "p") {
        const text = inline(child).trim();
        return text ? `${text}\n\n` : "";
      }
      if (t === "hr") return "---\n\n";
      if (t === "blockquote") return `${block(child).trim().split("\n").map((l) => `> ${l}`).join("\n")}\n\n`;
      if (t === "pre") return `\`\`\`\n${(child.textContent ?? "").replace(/\n+$/, "")}\n\`\`\`\n\n`;
      if (t === "ul" || t === "ol") return `${list(child, depth)}\n`;
      if (t === "ac:task-list") return `${tasks(child, depth)}\n`;
      if (t === "table") return `${table(child)}\n\n`;
      if (t === "ac:image") return `${image(child)}\n\n`;
      if (t === "ac:structured-macro") {
        const m = macroBlock(child);
        if (m !== undefined) return m;
        if (child.getAttribute("ac:name") === "status") return `${inline(child)}\n\n`;
        const body = bodyOf(child, "ac:rich-text-body") ?? bodyOf(child, "ac:plain-text-body");
        return body ? block(body, depth) : "";
      }
      return block(child, depth);
    })
    .join("");
}

/** A list item's text: its inline content (paragraphs inside joined), without its nested lists. */
const itemText = (li: Element) =>
  [...li.childNodes]
    .filter((c) => !(c instanceof Element && ["ul", "ol", "ac:task-list"].includes(tag(c))))
    .map((c) => inline(c))
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();

function list(node: Element, depth: number): string {
  const ordered = tag(node) === "ol";
  return [...node.children]
    .filter((li) => tag(li) === "li")
    .map((li, i) => {
      const nested = [...li.children].filter((c) => ["ul", "ol"].includes(tag(c)));
      return `${"  ".repeat(depth)}${ordered ? `${i + 1}.` : "-"} ${itemText(li)}\n${nested.map((n) => list(n, depth + 1)).join("")}`;
    })
    .join("");
}

function tasks(node: Element, depth: number): string {
  return [...node.children]
    .filter((t) => tag(t) === "ac:task")
    .map((t) => {
      const done = bodyOf(t, "ac:task-status")?.textContent?.trim() === "complete";
      const body = bodyOf(t, "ac:task-body");
      return `${"  ".repeat(depth)}- [${done ? "x" : " "}] ${body ? itemText(body) : ""}\n`;
    })
    .join("");
}

function table(node: Element): string {
  const rows = [...node.querySelectorAll("tr")].map((tr) => [...tr.children].map((c) => [...c.childNodes].map(inline).join(" ").replace(/\s+/g, " ").trim().replace(/\|/g, "\\|")));
  if (!rows.length) return "";
  const width = Math.max(...rows.map((r) => r.length));
  const line = (r: string[]) => `| ${Array.from({ length: width }, (_, i) => r[i] ?? "").join(" | ")} |`;
  return [line(rows[0]), `| ${Array.from({ length: width }, () => "---").join(" | ")} |`, ...rows.slice(1).map(line)].join("\n");
}
