/**
 * Rich text of "text" fields. A value is a small canonical HTML string (<b>, <i>, <u>,
 * <span style="color:#rrggbb">, <br>); in the Excel file it is stored as native rich text runs
 * (real bold / italic / colour), never as markdown.
 */
export type Run = { text: string; bold?: boolean; italic?: boolean; underline?: boolean; color?: string };

const escapeHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Normalizes any css colour (hex, rgb()) to #rrggbb. */
function hex(color: string): string | undefined {
  const c = color.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(c)) return c;
  if (/^#[0-9a-f]{3}$/.test(c)) return `#${c[1]}${c[1]}${c[2]}${c[2]}${c[3]}${c[3]}`;
  const m = c.match(/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)/);
  return m ? `#${[m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("")}` : undefined;
}

const BLOCKS = new Set(["DIV", "P", "LI"]);

/** HTML (canonical or produced by a contentEditable) → runs. */
export function htmlToRuns(html: string): Run[] {
  if (typeof DOMParser === "undefined") return [{ text: html }];
  const runs: Run[] = [];
  const walk = (node: Node, style: Omit<Run, "text">) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent ?? "";
      if (text) runs.push({ text, ...style });
      return;
    }
    if (!(node instanceof HTMLElement)) return;
    const next = { ...style };
    const css = node.style;
    if (/^(B|STRONG)$/.test(node.tagName) || /^(bold|[6-9]00)$/.test(css.fontWeight)) next.bold = true;
    if (/^(I|EM)$/.test(node.tagName) || css.fontStyle === "italic") next.italic = true;
    if (node.tagName === "U" || css.textDecorationLine.includes("underline") || css.textDecoration.includes("underline")) next.underline = true;
    const color = css.color || node.getAttribute("color");
    if (color) next.color = hex(color) ?? next.color;
    if (node.tagName === "BR") return void runs.push({ text: "\n" });
    if (BLOCKS.has(node.tagName) && runs.length && !runs[runs.length - 1].text.endsWith("\n")) runs.push({ text: "\n" });
    node.childNodes.forEach((child) => walk(child, next));
  };
  walk(new DOMParser().parseFromString(`<body>${html}</body>`, "text/html").body, {});
  // Trailing line break left by browsers, then merge equal neighbours.
  while (runs.length && /^\n*$/.test(runs[runs.length - 1].text)) runs.pop();
  const key = (r: Run) => `${r.bold}|${r.italic}|${r.underline}|${r.color}`;
  return runs.reduce<Run[]>((out, run) => {
    const last = out[out.length - 1];
    if (last && key(last) === key(run)) last.text += run.text;
    else out.push({ ...run });
    return out;
  }, []);
}

export function runsToHtml(runs: Run[]): string {
  return runs
    .map((run) => {
      let html = escapeHtml(run.text).replace(/\n/g, "<br>");
      if (run.color) html = `<span style="color:${run.color}">${html}</span>`;
      if (run.underline) html = `<u>${html}</u>`;
      if (run.italic) html = `<i>${html}</i>`;
      if (run.bold) html = `<b>${html}</b>`;
      return html;
    })
    .join("");
}

export const runsToText = (runs: Run[]) => runs.map((r) => r.text).join("");
export const isStyled = (runs: Run[]) => runs.some((r) => r.bold || r.italic || r.underline || r.color);

/** Canonical HTML of any text value (plain strings keep their line breaks). */
export const normalizeRich = (value: unknown) => (typeof DOMParser === "undefined" ? String(value) : runsToHtml(htmlToRuns(String(value))));
/** Plain text of a rich value (search, sort). */
export const richToText = (value: unknown) => runsToText(htmlToRuns(String(value ?? "")));
/** HTML of a plain string from a cell. */
export const plainToHtml = (text: string) => runsToHtml([{ text }]);
