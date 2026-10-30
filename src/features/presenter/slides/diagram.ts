/**
 * Diagrams of the slides (architecture, flows): boxes, groups of boxes and arrows, laid out by dagre
 * (the layered layout of Graphviz' dot, in the browser) and drawn by the templates in the deck's style.
 * A group is laid out inside first, then placed as one box; an arrow to a box inside a group is drawn
 * to the group. Units: cqw of a slide, before the template scales the whole to its box.
 */
import dagre from "@dagrejs/dagre";
import { z } from "zod";

export type DiagramNode = { id: string; title: string; description?: string; logo?: string; icon?: string; children?: DiagramNode[] };
export type DiagramEdge = { from: string; to: string; label?: string };
export type Direction = "LR" | "TB";

const node: z.ZodType<DiagramNode> = z.lazy(() =>
  z.object({
    id: z.string(),
    title: z.string(),
    description: z.string().optional(),
    logo: z.string().optional().describe("Logo: a public image URL (e.g. https://cdn.simpleicons.org/react) or data:image/svg+xml"),
    icon: z.string().optional().describe("Or a lucide icon name"),
    children: z.array(node).optional().describe("Boxes inside it: it becomes a group"),
  }),
);
export const diagramArgs = {
  nodes: z.array(node).min(1).max(12).describe("Boxes; a box with children is a group (e.g. the front end and its technologies)"),
  edges: z.array(z.object({ from: z.string(), to: z.string(), label: z.string().optional() })).max(16).default([]).describe("Arrows between boxes (ids)"),
  direction: z.enum(["LR", "TB"]).default("LR").describe("LR: left to right; TB: top to bottom"),
};

export type Placed = { node: DiagramNode; x: number; y: number; w: number; h: number; group: boolean; depth: number };
export type Route = { edge: DiagramEdge; points: { x: number; y: number }[]; label?: { x: number; y: number } };
export type Layout = { width: number; height: number; nodes: Placed[]; edges: Route[] };

/** Box sizes (cqw): a box grows with its description; a group adds its heading and margins. */
export const SIZE = { leaf: 18, head: 3.6, pad: 1.6, gap: 1.6, rank: 9 };
const leafHeight = (n: DiagramNode) => 4.4 + Math.min(3, Math.ceil((n.description?.length ?? 0) / 30)) * 1.45;

const contains = (n: DiagramNode, id: string): boolean => n.id === id || !!n.children?.some((c) => contains(c, id));

/** Lays out boxes and their arrows (those inside a group go with it). */
export function layoutDiagram(nodes: DiagramNode[], edges: DiagramEdge[], direction: Direction, depth = 0): Layout {
  // Groups first: their size is their inside's.
  const inner = new Map(nodes.filter((n) => n.children?.length).map((n) => [n.id, layoutDiagram(n.children!, edges.filter((e) => contains(n, e.from) && contains(n, e.to)), direction, depth + 1)]));
  const size = (n: DiagramNode) => {
    const l = inner.get(n.id);
    return l ? { w: Math.max(l.width + 2 * SIZE.pad, SIZE.leaf), h: l.height + SIZE.head + 2 * SIZE.pad } : { w: SIZE.leaf, h: leafHeight(n) };
  };
  // Arrows at this level: each end taken to the box of this level holding it.
  const at = (id: string) => nodes.find((n) => contains(n, id))?.id;
  const seen = new Set<string>();
  const level = edges.flatMap((e) => {
    const a = at(e.from);
    const b = at(e.to);
    const key = `${a}>${b}`;
    if (!a || !b || a === b || seen.has(key)) return [];
    seen.add(key);
    return [{ edge: e, a, b }];
  });
  const placed: Placed[] = [];
  const routes: Route[] = [];
  let width = 0;
  let height = 0;
  if (!level.length) {
    // No arrow: a column (left to right diagrams) or a row (top to bottom).
    let at = 0;
    for (const n of nodes) {
      const s = size(n);
      placed.push({ node: n, x: direction === "LR" ? 0 : at, y: direction === "LR" ? at : 0, ...s, group: inner.has(n.id), depth });
      at += (direction === "LR" ? s.h : s.w) + SIZE.gap;
      width = direction === "LR" ? Math.max(width, s.w) : at - SIZE.gap;
      height = direction === "LR" ? at - SIZE.gap : Math.max(height, s.h);
    }
    // Narrower boxes centered in the column / row.
    for (const p of placed) direction === "LR" ? (p.x = (width - p.w) / 2) : (p.y = (height - p.h) / 2);
  } else {
    const g = new dagre.graphlib.Graph({ multigraph: false });
    g.setGraph({ rankdir: direction, nodesep: SIZE.gap * 1.4, ranksep: SIZE.rank, edgesep: 1, marginx: 0, marginy: 0 });
    g.setDefaultEdgeLabel(() => ({}));
    for (const n of nodes) {
      const { w, h } = size(n);
      g.setNode(n.id, { width: w, height: h });
    }
    for (const { edge, a, b } of level) g.setEdge(a, b, edge.label ? { width: edge.label.length * 0.62 + 1.6, height: 1.8, labelpos: "c" } : {});
    dagre.layout(g);
    for (const n of nodes) {
      const v = g.node(n.id);
      placed.push({ node: n, x: v.x - v.width / 2, y: v.y - v.height / 2, w: v.width, h: v.height, group: inner.has(n.id), depth });
    }
    for (const { edge, a, b } of level) {
      const e = g.edge(a, b) as { points: { x: number; y: number }[]; x?: number; y?: number };
      routes.push({ edge, points: e.points, label: edge.label && e.x !== undefined ? { x: e.x, y: e.y! } : undefined });
    }
    width = g.graph().width ?? 0;
    height = g.graph().height ?? 0;
  }
  // The inside of each group, moved to its place.
  const nested = placed.flatMap((p) => {
    const l = inner.get(p.node.id);
    if (!l) return [];
    const dx = p.x + (p.w - l.width) / 2;
    const dy = p.y + SIZE.head + SIZE.pad;
    return [{ nodes: l.nodes.map((c) => ({ ...c, x: c.x + dx, y: c.y + dy })), edges: l.edges.map((r) => ({ ...r, points: r.points.map((q) => ({ x: q.x + dx, y: q.y + dy })), label: r.label && { x: r.label.x + dx, y: r.label.y + dy } })) }];
  });
  return { width, height, nodes: [...placed, ...nested.flatMap((n) => n.nodes)], edges: [...routes, ...nested.flatMap((n) => n.edges)] };
}

/** An arrow with right angles (rounded corners), leaving and arriving along the diagram's direction. */
export function elbow(points: { x: number; y: number }[], direction: Direction, r = 1.2) {
  const a = points[0];
  const b = points[points.length - 1];
  if (direction === "LR") {
    const mx = (a.x + b.x) / 2;
    const dy = Math.sign(b.y - a.y);
    const k = Math.min(r, Math.abs(b.y - a.y) / 2, Math.abs(mx - a.x));
    if (!dy || k < 0.05) return `M${a.x} ${a.y} L${b.x} ${b.y}`;
    return `M${a.x} ${a.y} H${mx - k} Q${mx} ${a.y} ${mx} ${a.y + dy * k} V${b.y - dy * k} Q${mx} ${b.y} ${mx + k} ${b.y} H${b.x}`;
  }
  const my = (a.y + b.y) / 2;
  const dx = Math.sign(b.x - a.x);
  const k = Math.min(r, Math.abs(b.x - a.x) / 2, Math.abs(my - a.y));
  if (!dx || k < 0.05) return `M${a.x} ${a.y} L${b.x} ${b.y}`;
  return `M${a.x} ${a.y} V${my - k} Q${a.x} ${my} ${a.x + dx * k} ${my} H${b.x - dx * k} Q${b.x} ${my} ${b.x} ${my + k} V${b.y}`;
}

/** The corners of the same arrow, without rounding (PowerPoint). */
export function elbowPoints(points: { x: number; y: number }[], direction: Direction) {
  const a = points[0];
  const b = points[points.length - 1];
  if (direction === "LR") {
    const mx = (a.x + b.x) / 2;
    return [a, { x: mx, y: a.y }, { x: mx, y: b.y }, b];
  }
  const my = (a.y + b.y) / 2;
  return [a, { x: a.x, y: my }, { x: b.x, y: my }, b];
}
