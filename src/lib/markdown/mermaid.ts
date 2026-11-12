/**
 * The Mermaid flowcharts of the markdown (```mermaid), the subset the models write well: `flowchart LR`
 * (or TD), `id[Label]`, `id(Label)`, `id[(Database)]`, `subgraph id[Label] … end`, `a --> b`,
 * `a -->|label| b`, `a -- label --> b`, `-.->` (dashed), `---` (no arrow), chains `a --> b --> c`,
 * colors by palette names (`:::blue`, `class a,b blue`), an icon before a label (`icon:database Sales`
 * or `fa:fa-database Sales`) and a detail after `<br/>`. Drawn by components/diagram.tsx.
 */

export type GNode = { id: string; label: string; icon?: string; color?: string; detail?: string; parent?: string };
export type GEdge = { from: string; to: string; label?: string; dashed?: boolean; line?: boolean };
export type Graph = { direction: "LR" | "TB"; nodes: GNode[]; groups: GNode[]; edges: GEdge[] };

/** A flowchart (shapes are kept as icons: a cylinder is a database). */
export function parseMermaid(text: string): Graph {
  const g: Graph = { direction: "LR", nodes: [], groups: [], edges: [] };
  const stack: string[] = [];
  /** `"icon:database Sales<br/>import mode"` → icon, label and the detail under it. */
  const label = (raw: string): Partial<GNode> & { label: string } => {
    const [s, ...more] = raw.replace(/^["']|["']$/g, "").trim().split(/<br\s*\/?>/i);
    const m = s.match(/^(?:icon:|fa:fa-)([\w-]+)\s+(.*)$/);
    const detail = more.join(" ").trim() || undefined;
    return m ? { icon: m[1], label: m[2].trim(), detail } : { label: s.trim(), detail };
  };
  /** `id`, `id[Label]`, `id(Label)`, `id[(Label)]`, `id:::color` — declared once, refined after. */
  const ref = (token: string) => {
    const m = token.trim().match(/^([\w-]+)\s*(\[\(|\(\[|\[|\(|\{)?\s*(.*?)\s*(?:\)\]|\]\)|\]|\)|\})?\s*(?::::(\w+))?$/);
    if (!m) return undefined;
    const [, id, open, inner, cls] = m;
    // An id that names a group is that group (an arrow to it goes to the whole group): never a second box with the same id.
    const group = g.groups.find((x) => x.id === id);
    let n = g.nodes.find((x) => x.id === id) ?? group;
    if (!n) g.nodes.push((n = { id, label: id, parent: stack.at(-1) }));
    if (open && inner) Object.assign(n, label(inner), open === "[(" && !label(inner).icon ? { icon: "database" } : {});
    if (cls) n.color = cls;
    return n;
  };
  for (const raw of text.split("\n")) {
    const line = raw.trim();
    if (!line || line.startsWith("%%")) continue;
    let m = line.match(/^(?:flowchart|graph)\s+(LR|RL|TB|TD|BT)/i);
    if (m) {
      g.direction = /LR|RL/i.test(m[1]) ? "LR" : "TB";
      continue;
    }
    m = line.match(/^subgraph\s+([\w-]+)(?:\s*\[(.+)\])?/);
    if (m) {
      // A group named like a box already drawn gets its own id (both must stay).
      const id = g.nodes.some((x) => x.id === m![1]) || g.groups.some((x) => x.id === m![1]) ? `${m[1]}~group${g.groups.length}` : m[1];
      g.groups.push({ id, label: m[2] ? label(m[2]).label : m[1], parent: stack.at(-1) });
      stack.push(id);
      continue;
    }
    if (line === "end") {
      stack.pop();
      continue;
    }
    // Colors: the class names of the palette (`classDef blue fill:#eff6ff` is drawn with its own colors anyway).
    if (line.startsWith("classDef ") || line.startsWith("style ")) continue;
    m = line.match(/^class\s+([\w,-]+)\s+(\w+)/);
    if (m) {
      for (const id of m[1].split(",")) {
        const n = g.nodes.find((x) => x.id === id) ?? g.groups.find((x) => x.id === id);
        if (n) n.color = m[2];
      }
      continue;
    }
    // A chain: a -->|x| b -.-> c (`a -- x --> b` is the same as `a -->|x| b`; `---` a line without arrow).
    const parts = line.replace(/\s--\s+([^-|>][^>]*?)\s+-->/g, " -->|$1| ").split(/\s*(-->|-\.->|==>|---)\s*(?:\|([^|]*)\|)?\s*/);
    if (parts.length >= 4) {
      for (let i = 0; i + 3 < parts.length; i += 3) {
        const a = ref(parts[i]);
        const b = ref(parts[i + 3]);
        if (a && b) g.edges.push({ from: a.id, to: b.id, label: parts[i + 2]?.trim() || undefined, dashed: parts[i + 1] === "-.->", line: parts[i + 1] === "---" });
      }
      continue;
    }
    ref(line);
  }
  return g;
}

