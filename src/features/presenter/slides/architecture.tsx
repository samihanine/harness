import { useId, useMemo } from "react";
import { z } from "zod";
import { pptxImage } from "./images";
import { Icon } from "./picture";
import { Frame, M, STAGE, deep, defineTemplate, ease, faint, hex, mix, pptxPath, pptxTitle, tint } from "./kit";
import { diagramArgs, elbow, elbowPoints, layoutDiagram, type DiagramNode, type Placed } from "./diagram";

/** A box's logo (an image) or icon. */
function Logo({ node, size }: { node: DiagramNode; size: number }) {
  const box = { width: `${size}cqw`, height: `${size}cqw` };
  if (node.logo) return <img src={node.logo} alt="" className="shrink-0 object-contain" style={box} />;
  return <Icon name={node.icon ?? "box"} className="shrink-0" style={box} />;
}

/** Simplified logos drawn here (no public link to rely on). */
const svg = (body: string) => `data:image/svg+xml;utf8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">${body}</svg>`)}`;
const POWER_BI = svg('<rect x="13.5" y="2" width="6" height="20" rx="1.6" fill="#F2C811"/><rect x="8.8" y="7" width="6" height="15" rx="1.6" fill="#E3A90C"/><rect x="4.1" y="12" width="6" height="10" rx="1.6" fill="#C38A06"/>');
const FABRIC = svg('<path d="M5 7.5 15 3l5 2.3-10 4.5z" fill="#1AAE9F"/><path d="M5 13.5 15 9l5 2.3-10 4.5z" fill="#3DC48E"/><path d="M5 7.5l5 2.3V21l-5-2.3z" fill="#0F7C8C"/>');
const simple = (slug: string) => `https://cdn.simpleicons.org/${slug}`;

/** The scale of a diagram in the content box (cqw per unit of the layout). */
const scaleOf = (width: number, height: number) => Math.min(1.25, STAGE.w / width, (STAGE.h + 2) / height);

/**
 * Architecture (or any boxes-and-arrows diagram): boxes, groups of boxes (a dashed frame with a tab)
 * and arrows laid out by dagre (diagram.ts), drawn as white cards with right-angled arrows along
 * which data flows.
 */
export const architecture = defineTemplate({
  id: "architecture",
  description:
    "A diagram of boxes and arrows, laid out automatically (an architecture, a flow): nodes (id, title, description, logo URL or lucide icon; children: boxes inside it, e.g. the front end and its technologies), edges (from, to, optional label), direction LR (left to right) or TB (top to bottom).",
  args: z.object({ title: z.string(), ...diagramArgs }),
  example: {
    title: "Architecture",
    direction: "LR",
    nodes: [
      {
        id: "front",
        title: "Front end",
        description: "Tout dans le navigateur",
        icon: "monitor",
        children: [
          { id: "react", title: "React", description: "Écrans, assistants et slides", logo: simple("react") },
          { id: "ts", title: "TypeScript", description: "Le harness des agents, typé", logo: simple("typescript") },
          { id: "bun", title: "Bun", description: "Dev, build et serveur de jetons", logo: simple("bun") },
        ],
      },
      { id: "ai", title: "API IA", description: "LLM privé : texte en entrée et en sortie", icon: "sparkles" },
      { id: "pbi", title: "Microsoft Power BI", description: "Rapports, embed et requêtes DAX", logo: POWER_BI },
      { id: "fabric", title: "Microsoft Fabric", description: "Définitions PBIR et TMDL", logo: FABRIC },
      { id: "confluence", title: "Confluence", description: "Pages de documentation", logo: simple("confluence") },
    ],
    edges: [
      { from: "front", to: "ai", label: "prompts" },
      { from: "front", to: "pbi", label: "REST · DAX" },
      { from: "front", to: "fabric", label: "définitions" },
      { from: "front", to: "confluence", label: "REST" },
    ],
  },
  Component: ({ args }) => {
    const layout = useMemo(() => layoutDiagram(args.nodes, args.edges, args.direction), [args.nodes, args.edges, args.direction]);
    const id = useId().replace(/:/g, "");
    const s = scaleOf(layout.width, layout.height);
    const at = (p: Placed) => ({ left: `${p.x * s}cqw`, top: `${p.y * s}cqw`, width: `${p.w * s}cqw`, height: `${p.h * s}cqw` });
    return (
      <Frame title={args.title}>
        <div className="flex h-full items-center justify-center">
          <div className="relative" style={{ width: `${layout.width * s}cqw`, height: `${layout.height * s}cqw` }}>
            {/* Groups under the arrows, boxes over them. */}
            {layout.nodes
              .filter((p) => p.group)
              .map((p) => (
                <M.div key={p.node.id} initial={{ opacity: 0 }} animate={{ opacity: 1 }} transition={{ duration: 0.6 }} className="absolute rounded-[1cqw] border-dashed" style={{ ...at(p), borderWidth: `${0.16 * s}cqw`, borderColor: "var(--s-primary)", background: tint(6) }}>
                  <span className="absolute flex items-center gap-[0.6cqw] rounded-full px-[1cqw] py-[0.3cqw] font-bold text-white" style={{ left: `${1.2 * s}cqw`, top: `${-1.1 * s}cqw`, fontSize: `${1.15 * s}cqw`, background: deep }}>
                    {(p.node.logo || p.node.icon) && <Logo node={p.node} size={1.4 * s} />}
                    {p.node.title}
                  </span>
                  {p.node.description && (
                    <span className="absolute right-[1.2cqw] opacity-60" style={{ top: `${0.9 * s}cqw`, fontSize: `${0.95 * s}cqw`, left: `${1.6 * s}cqw` }}>
                      {p.node.description}
                    </span>
                  )}
                </M.div>
              ))}
            <svg viewBox={`0 0 ${layout.width} ${layout.height}`} className="pointer-events-none absolute inset-0 size-full overflow-visible" fill="none">
              <defs>
                <marker id={`${id}-arrow`} viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse">
                  <path d="M0 1 L9 5 L0 9 Z" style={{ fill: faint(45) }} />
                </marker>
              </defs>
              {layout.edges.map((r, i) => (
                <g key={i}>
                  <M.path d={elbow(r.points, args.direction)} style={{ stroke: faint(25) }} strokeWidth={0.16} markerEnd={`url(#${id}-arrow)`} initial={{ pathLength: 0 }} animate={{ pathLength: 1 }} transition={{ delay: 0.8 + i * 0.12, duration: 0.7, ease }} />
                  {/* Data flowing along the arrow. */}
                  <M.path d={elbow(r.points, args.direction)} style={{ stroke: "var(--s-primary)" }} strokeWidth={0.22} strokeDasharray="0.8 2.2" initial={{ strokeDashoffset: 0, opacity: 0 }} animate={{ strokeDashoffset: -12, opacity: 1 }} transition={{ strokeDashoffset: { repeat: Infinity, duration: 2, ease: "linear" }, opacity: { delay: 1.5 } }} />
                </g>
              ))}
            </svg>
            {layout.nodes
              .filter((p) => !p.group)
              .map((p, i) => (
                <M.div key={p.node.id} initial={{ opacity: 0, y: "0.6cqw" }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.3 + i * 0.08, duration: 0.5, ease }} whileHover={{ y: "-0.3cqw" }} className="absolute flex flex-col gap-[0.4cqw] rounded-[0.8cqw] border shadow-[0_0.25cqw_0.8cqw_rgba(0,0,0,0.08)]" style={{ ...at(p), padding: `${1 * s}cqw ${1.1 * s}cqw`, borderColor: faint(12), background: "var(--s-bg)" }}>
                  <span className="flex items-center gap-[0.8cqw]">
                    <span className="grid shrink-0 place-items-center rounded-[0.5cqw]" style={{ width: `${2.6 * s}cqw`, height: `${2.6 * s}cqw`, background: faint(5), color: deep }}>
                      <Logo node={p.node} size={1.7 * s} />
                    </span>
                    <span className="leading-tight font-bold" style={{ fontSize: `${1.25 * s}cqw` }}>
                      {p.node.title}
                    </span>
                  </span>
                  {p.node.description && (
                    <span className="leading-snug opacity-70" style={{ fontSize: `${0.95 * s}cqw` }}>
                      {p.node.description}
                    </span>
                  )}
                </M.div>
              ))}
            {layout.edges.map(
              (r, i) =>
                r.label &&
                r.edge.label && (
                  <M.span key={`l${i}`} initial={{ opacity: 0 }} animate={{ opacity: 1 }} transition={{ delay: 1.4 + i * 0.1 }} className="absolute -translate-x-1/2 -translate-y-1/2 rounded-full px-[0.6cqw] py-[0.1cqw] whitespace-nowrap" style={{ left: `${r.label.x * s}cqw`, top: `${r.label.y * s}cqw`, fontSize: `${0.85 * s}cqw`, background: "var(--s-bg)", border: `1px solid ${faint(15)}` }}>
                    {r.edge.label}
                  </M.span>
                ),
            )}
          </div>
        </div>
      </Frame>
    );
  },
  pptx: (slide, args, ctx) => {
    pptxTitle(slide, args.title, ctx);
    const { theme } = ctx;
    const layout = layoutDiagram(args.nodes, args.edges, args.direction);
    // The same scale as on screen: 1 cqw is 1/100 of the slide's width.
    const k = (scaleOf(layout.width, layout.height) * 13.333) / 100;
    const x0 = 0.8 + (11.7 - layout.width * k) / 2;
    const y0 = 2.1 + (4.8 - layout.height * k) / 2;
    const at = (p: { x: number; y: number }) => ({ x: x0 + p.x * k, y: y0 + p.y * k });
    const picture = (n: DiagramNode, box: { x: number; y: number; w: number; h: number }, white = false) => {
      const image = pptxImage(n.logo ?? `${white ? "icon-white" : "icon"}:${n.icon?.trim() || "box"}`, box);
      if (image) slide.addImage(image);
    };
    const text = hex(theme.text);
    for (const p of layout.nodes.filter((p) => p.group)) {
      const { x, y } = at(p);
      slide.addShape("roundRect", { x, y, w: p.w * k, h: p.h * k, fill: { color: mix(theme.primary, theme.background, 0.06) }, line: { color: hex(theme.primary), width: 1.25, dashType: "dash" }, rectRadius: 0.1 });
      const tab = Math.max(1.2, p.node.title.length * 0.1 + 0.6);
      slide.addShape("roundRect", { x: x + 0.15, y: y - 0.15, w: tab, h: 0.3, fill: { color: mix(theme.primary, theme.text, 0.62) }, line: { type: "none" } as never, rectRadius: 0.15 });
      if (p.node.logo || p.node.icon) picture(p.node, { x: x + 0.25, y: y - 0.1, w: 0.2, h: 0.2 }, true);
      slide.addText(p.node.title, { x: x + 0.5, y: y - 0.15, w: tab - 0.4, h: 0.3, fontSize: 10, bold: true, color: "FFFFFF", fontFace: theme.fontText, valign: "middle" });
      if (p.node.description) slide.addText(p.node.description, { x: x + 0.2, y: y + 0.15, w: p.w * k - 0.4, h: 0.3, fontSize: 9, color: text, transparency: 40, fontFace: theme.fontText });
    }
    layout.edges.forEach((r) => {
      const points = elbowPoints(r.points, args.direction).map(at);
      pptxPath(slide, points, { color: mix(theme.text, theme.background, 0.3), width: 1, arrow: true });
      pptxPath(slide, points, { color: hex(theme.primary), width: 1.75, dash: "sysDot" });
      if (r.label && r.edge.label) {
        const c = at(r.label);
        const w = r.edge.label.length * 0.075 + 0.25;
        slide.addShape("roundRect", { x: c.x - w / 2, y: c.y - 0.12, w, h: 0.24, fill: { color: hex(theme.background) }, line: { color: mix(theme.text, theme.background, 0.15), width: 0.5 }, rectRadius: 0.12 });
        slide.addText(r.edge.label, { x: c.x - w / 2, y: c.y - 0.12, w, h: 0.24, fontSize: 8, align: "center", valign: "middle", color: text, fontFace: theme.fontText });
      }
    });
    for (const p of layout.nodes.filter((p) => !p.group)) {
      const { x, y } = at(p);
      const w = p.w * k;
      slide.addShape("roundRect", { x, y, w, h: p.h * k, fill: { color: hex(theme.background) }, line: { color: mix(theme.text, theme.background, 0.12), width: 0.75 }, rectRadius: 0.08, shadow: { type: "outer", blur: 4, offset: 1, angle: 90, color: "000000", opacity: 0.12 } });
      slide.addShape("roundRect", { x: x + 0.13, y: y + 0.13, w: 0.35, h: 0.35, fill: { color: mix(theme.text, theme.background, 0.05) }, line: { type: "none" } as never, rectRadius: 0.15 });
      picture(p.node, { x: x + 0.18, y: y + 0.18, w: 0.25, h: 0.25 });
      slide.addText(p.node.title, { x: x + 0.55, y: y + 0.13, w: w - 0.65, h: 0.35, fontSize: 11, bold: true, color: text, fontFace: theme.fontText, valign: "middle" });
      if (p.node.description) slide.addText(p.node.description, { x: x + 0.1, y: y + 0.52, w: w - 0.2, h: p.h * k - 0.6, fontSize: 9, color: text, transparency: 30, fontFace: theme.fontText, valign: "top" });
    }
  },
});
