/**
 * The one markdown renderer of the app (answers, guides, analyses, Confluence pages, slides' notes):
 * GitHub markdown plus the constructs of lib/markdown/syntax.ts (mermaid diagrams, colored code,
 * panels, statuses, icons, task lists), the same that are written to Confluence.
 */
import { Children, isValidElement, type ReactElement, type ReactNode } from "react";
import ReactMarkdown, { type Components } from "react-markdown";
import remarkGfm from "remark-gfm";
import { ChevronRightIcon, CircleAlertIcon, InfoIcon, LightbulbIcon, TriangleAlertIcon } from "lucide-react";
import { Mermaid } from "./diagram";
import { ErrorBoundary } from "./error-boundary";
import { iconOf } from "./icon";
import { highlight } from "@/lib/markdown/highlight";
import { panelOf, safeUrl, statusOf, withPanels, type PanelKind } from "@/lib/markdown/syntax";
import { cn } from "@/lib/utils";

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

const PANELS: Record<PanelKind, { icon: typeof InfoIcon; className: string; bar: string }> = {
  NOTE: { icon: InfoIcon, className: "bg-blue-500/8 [&_svg]:text-blue-600", bar: "bg-blue-500" },
  TIP: { icon: LightbulbIcon, className: "bg-emerald-500/8 [&_svg]:text-emerald-600", bar: "bg-emerald-500" },
  WARNING: { icon: TriangleAlertIcon, className: "bg-amber-500/10 [&_svg]:text-amber-600", bar: "bg-amber-500" },
  IMPORTANT: { icon: CircleAlertIcon, className: "bg-violet-500/8 [&_svg]:text-violet-600", bar: "bg-violet-500" },
  EXPAND: { icon: ChevronRightIcon, className: "border bg-muted/40", bar: "" },
};
const STATUS: Record<string, string> = { green: "bg-emerald-600", blue: "bg-blue-600", red: "bg-red-600", yellow: "bg-amber-500", grey: "bg-zinc-500", gray: "bg-zinc-500", purple: "bg-violet-600" };

function CodeBlock({ lang, code }: { lang?: string; code: string }) {
  return (
    <div className="overflow-hidden rounded-lg border bg-zinc-50 dark:bg-zinc-900">
      {lang && <div className="flex h-7 items-center border-b px-3 text-[11px] font-medium tracking-wide text-muted-foreground uppercase">{lang}</div>}
      <pre className="!m-0 overflow-x-auto !rounded-none !bg-transparent p-3 text-[12.5px] leading-relaxed">
        <code className="!bg-transparent !p-0">{highlight(code, lang)}</code>
      </pre>
    </div>
  );
}

/** A fenced block: a diagram, a panel or code. */
function Fence({ lang, code }: { lang: string; code: string }) {
  if (lang === "mermaid")
    return (
      <ErrorBoundary resetKey={code} fallback={(e) => <CodeBlock lang={`mermaid — not drawn: ${e.message.slice(0, 80)}`} code={code} />}>
        <Mermaid code={code} />
      </ErrorBoundary>
    );
  const panel = panelOf(lang, code);
  if (panel) {
    const { icon: Icon, className, bar } = PANELS[panel.kind];
    if (panel.kind === "EXPAND")
      return (
        <details className={cn("rounded-lg border px-3 py-2", className)}>
          <summary className="cursor-pointer text-[13.5px] font-medium">{panel.title || "Details"}</summary>
          <div className="pt-2">
            <Markdown>{panel.body}</Markdown>
          </div>
        </details>
      );
    return (
      <div className={cn("flex overflow-hidden rounded-lg", className)}>
        <span className={cn("w-1 shrink-0", bar)} />
        <div className="flex min-w-0 gap-2.5 px-3 py-2.5">
          <Icon className="mt-0.5 size-4 shrink-0" />
          <div className="flex min-w-0 flex-col gap-1">
            {panel.title && <b>{panel.title}</b>}
            <Markdown>{panel.body}</Markdown>
          </div>
        </div>
      </div>
    );
  }
  return <CodeBlock lang={lang} code={code} />;
}

const components: Components = {
  pre: ({ children }) => {
    const code = Children.toArray(children).find(isValidElement) as ReactElement<{ className?: string; children?: ReactNode }> | undefined;
    return <Fence lang={code?.props.className?.match(/language-([\w-]+)/)?.[1] ?? ""} code={String(code?.props.children ?? "").replace(/\n$/, "")} />;
  },
  code: ({ children, className }) => {
    const status = !className && statusOf(String(children));
    if (status) return <span className={cn("inline-flex h-5 items-center rounded px-1.5 align-middle text-[10.5px] font-bold tracking-wide text-white uppercase", STATUS[status.color] ?? STATUS.grey)}>{status.text}</span>;
    return <code className={className}>{children}</code>;
  },
  img: ({ src, alt }) => {
    const icon = String(src ?? "").match(/^icon:([\w-]+)$/)?.[1];
    if (icon) {
      const Icon = iconOf(icon);
      return <Icon className="inline size-[1.15em] -translate-y-px align-middle text-primary" aria-label={alt || icon} />;
    }
    // A picture attached to a Confluence page: shown by its name (the app does not download attachments).
    if (String(src).startsWith("attachment:")) return <span className="rounded bg-muted px-1.5 text-[12px] text-muted-foreground">🖼 {alt || String(src).slice(11)}</span>;
    return <img src={src as string} alt={alt ?? ""} className="max-w-full rounded" />;
  },
  // Links leave the app in a new tab, without handing it over (no opener, no referrer).
  a: ({ href, children }) => (
    <a href={href} target={href?.startsWith("#") ? undefined : "_blank"} rel="noopener noreferrer">
      {children}
    </a>
  ),
  table: ({ children }) => (
    <div className="max-w-full overflow-x-auto">
      <table>{children}</table>
    </div>
  ),
};

export const Markdown = ({ children, className }: { children: string; className?: string }) => (
  // A text the renderer cannot draw is shown as it is, never a blank page.
  <ErrorBoundary resetKey={children} fallback={() => <pre className="overflow-auto rounded-lg bg-muted p-3 text-[12.5px] whitespace-pre-wrap">{children}</pre>}>
  <div className={cn("markdown [&_.contains-task-list]:list-none [&_.contains-task-list]:pl-1 [&_.task-list-item_input]:mr-1.5 [&_.task-list-item_input]:align-middle", className)}>
    <ReactMarkdown remarkPlugins={[remarkGfm]} components={components} urlTransform={(url, key) => safeUrl(url, key === "src")}>
      {withPanels(cleanMarkdown(children))}
    </ReactMarkdown>
  </div>
  </ErrorBoundary>
);
