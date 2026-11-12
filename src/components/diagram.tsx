/**
 * A ```mermaid block drawn as one SVG (boxes with an icon, colored groups, arrows with labels), laid
 * out by lib/markdown/layout.ts (dagre). Pure SVG with fixed colors (icons included): the same drawing
 * is attached to the Confluence page as a picture (`diagramPng`).
 */
import { createElement, useId } from "react";
import { iconOf } from "./icon";
import { renderHtml } from "@/lib/markdown/render";
import { elbow, layoutDiagram, type DiagramNode } from "@/lib/markdown/layout";
import { parseMermaid, type Graph } from "@/lib/markdown/mermaid";

/** Fixed colors (the SVG leaves the app as an image: no theme variables in it). */
const PALETTE: Record<string, { fill: string; stroke: string; text: string }> = {
  gray: { fill: "#f4f4f5", stroke: "#a1a1aa", text: "#3f3f46" },
  blue: { fill: "#eff6ff", stroke: "#3b82f6", text: "#1d4ed8" },
  violet: { fill: "#f5f3ff", stroke: "#8b5cf6", text: "#6d28d9" },
  teal: { fill: "#f0fdfa", stroke: "#14b8a6", text: "#0f766e" },
  green: { fill: "#ecfdf5", stroke: "#10b981", text: "#047857" },
  amber: { fill: "#fffbeb", stroke: "#f59e0b", text: "#b45309" },
  orange: { fill: "#fff7ed", stroke: "#f97316", text: "#c2410c" },
  red: { fill: "#fef2f2", stroke: "#ef4444", text: "#b91c1c" },
  pink: { fill: "#fdf2f8", stroke: "#ec4899", text: "#be185d" },
};
const paint = (c?: string) => PALETTE[c ?? ""] ?? PALETTE.gray;

/* ---------------------------------- Drawing --------------------------------- */

/** Pixels per unit of the layout. */
const S = 10;

/** The graph as the layout's tree: a group holds its boxes and groups; arrows into a group go to it. */
function toTree(g: Graph): DiagramNode[] {
  const of = (parent?: string): DiagramNode[] => [
    ...g.groups.filter((x) => x.parent === parent).map((x) => ({ id: x.id, title: x.label, children: of(x.id) })),
    ...g.nodes.filter((x) => (parent ? x.parent === parent : !x.parent || !g.groups.some((gr) => gr.id === x.parent))).map((x) => ({ id: x.id, title: x.label, description: x.detail, icon: x.icon })),
  ];
  return of(undefined);
}

/** A description on lines of about 26 characters (the box's height allows for them). */
const lines = (text: string) => text.split(" ").reduce<string[]>((out, w) => (out.length && (out.at(-1)! + " " + w).length <= 26 ? ((out[out.length - 1] += " " + w), out) : [...out, w]), []);

const iconSvg = (name: string, x: number, y: number, size: number, color: string) => createElement(iconOf(name), { x, y, width: size, height: size, color, strokeWidth: 2 });

/**
 * The diagram as an SVG element (or the parse error in its place), laid out by the slides' engine
 * (`layoutDiagram`: groups laid out inside first, then placed as one box) — one engine for slides and pages.
 */
export function DiagramSvg({ graph }: { graph: Graph }) {
  // Its own arrowhead: an id shared with a diagram in a hidden tab would not draw.
  const arrow = `dg-arrow-${useId().replace(/[^\w-]/g, "")}`;
  let l: ReturnType<typeof layoutDiagram>;
  try {
    if (!graph.nodes.length) throw new Error("No box in the diagram.");
    l = layoutDiagram(toTree(graph), graph.edges.map((e) => ({ from: e.from, to: e.to, label: e.label })), graph.direction);
  } catch (e) {
    return <pre className="rounded-md bg-red-500/10 p-2 text-[12px] text-destructive">{String(e)}</pre>;
  }
  const color = new Map([...graph.groups, ...graph.nodes].map((n) => [n.id, n.color]));
  const dashed = new Set(graph.edges.filter((e) => e.dashed).map((e) => `${e.from}>${e.to}`));
  const plain = new Set(graph.edges.filter((e) => e.line).map((e) => `${e.from}>${e.to}`));
  const pad = 8;
  return (
    <svg viewBox={`${-pad} ${-pad} ${l.width * S + 2 * pad} ${l.height * S + 2 * pad}`} width={l.width * S + 2 * pad} className="h-auto max-w-full" fontFamily="'Plus Jakarta Sans Variable', system-ui, sans-serif" role="img">
      <defs>
        <marker id={arrow} viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
          <path d="M0,0 L10,5 L0,10 z" fill="#71717a" />
        </marker>
      </defs>
      {[...l.nodes].sort((a, b) => a.depth - b.depth).filter((p) => p.group).map((p) => {
        const c = paint(color.get(p.node.id));
        return (
          <g key={p.node.id}>
            <rect x={p.x * S} y={p.y * S} width={p.w * S} height={p.h * S} rx={12} fill={c.fill} stroke={c.stroke} strokeOpacity={0.5} strokeDasharray={color.get(p.node.id) ? undefined : "4 3"} />
            <text x={p.x * S + 12} y={p.y * S + 21} fontSize={11} fontWeight={700} fill={c.text} letterSpacing={0.5}>
              {p.node.title.toUpperCase()}
            </text>
          </g>
        );
      })}
      {l.edges.map((r: (typeof l.edges)[number], i: number) => {
        const pts = r.points.map((q: { x: number; y: number }) => ({ x: q.x * S, y: q.y * S }));
        return (
          <g key={i}>
            <path d={elbow(pts, graph.direction, 10)} fill="none" stroke="#71717a" strokeWidth={1.4} strokeDasharray={dashed.has(`${r.edge.from}>${r.edge.to}`) ? "5 4" : undefined} markerEnd={plain.has(`${r.edge.from}>${r.edge.to}`) ? undefined : `url(#${arrow})`} />
            {r.label && (
              <g>
                <rect x={r.label.x * S - r.edge.label!.length * 3.1 - 6} y={r.label.y * S - 9} width={r.edge.label!.length * 6.2 + 12} height={18} rx={9} fill="#ffffff" stroke="#e4e4e7" />
                <text x={r.label.x * S} y={r.label.y * S + 3.5} fontSize={10.5} textAnchor="middle" fill="#52525b">
                  {r.edge.label}
                </text>
              </g>
            )}
          </g>
        );
      })}
      {l.nodes.filter((p) => !p.group).map((p) => {
        const c = paint(color.get(p.node.id));
        const [x, y, w, h] = [p.x * S, p.y * S, p.w * S, p.h * S];
        const cy = y + h / 2;
        const tx = x + (p.node.icon ? 46 : 12);
        const desc = p.node.description ? lines(p.node.description) : [];
        const top = cy - (desc.length * 12) / 2 + 4;
        return (
          <g key={p.node.id}>
            <rect x={x} y={y} width={w} height={h} rx={9} fill="#ffffff" stroke={c.stroke} strokeWidth={1.3} />
            {p.node.icon && (
              <>
                <rect x={x + 8} y={cy - 15} width={30} height={30} rx={8} fill={c.fill} />
                {iconSvg(p.node.icon, x + 14, cy - 9, 18, c.text)}
              </>
            )}
            <text x={tx} y={top} fontSize={12.5} fontWeight={600} fill="#18181b">
              {p.node.title}
            </text>
            {desc.map((d, j) => (
              <text key={j} x={tx} y={top + 13 + j * 12} fontSize={10} fill="#71717a">
                {d}
              </text>
            ))}
          </g>
        );
      })}
    </svg>
  );
}

/** A mermaid block drawn (its parse error said in place of it). */
export function Mermaid({ code }: { code: string }) {
  return (
    <figure className="flex justify-center overflow-x-auto rounded-lg border bg-white p-3">
      <DiagramSvg graph={parseMermaid(code)} />
    </figure>
  );
}

/** The SVG of a mermaid block as text (white background), for the picture attached to Confluence. */
export function diagramSvgText(code: string) {
  const svg = renderHtml(<DiagramSvg graph={parseMermaid(code)} />);
  return svg.replace("<svg ", '<svg xmlns="http://www.w3.org/2000/svg" style="background:#ffffff" ').replace(/ class="[^"]*"/, "");
}

/** An SVG text as a PNG (twice its size: sharp on screens), through a canvas. */
export async function svgToPng(svg: string): Promise<Blob> {
  const img = new Image();
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  await img.decode();
  const k = 2;
  const canvas = Object.assign(document.createElement("canvas"), { width: img.naturalWidth * k, height: img.naturalHeight * k });
  const ctx = canvas.getContext("2d")!;
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  return new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("The diagram could not be drawn."))), "image/png"));
}

/** A lucide icon as a PNG (for the icons of a page written to Confluence). */
export const iconPng = (name: string, color = "#2563eb") =>
  svgToPng(renderHtml(createElement(iconOf(name), { width: 32, height: 32, color, strokeWidth: 2, xmlns: "http://www.w3.org/2000/svg" })));
