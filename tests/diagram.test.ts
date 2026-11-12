import { describe, expect, test } from "bun:test";
import { createElement } from "react";
import { DiagramSvg } from "@/components/diagram";
import { renderHtml } from "@/lib/markdown/render";
import { parseMermaid } from "@/lib/markdown/mermaid";

/** Diagrams the way models write them, good and bad: each must draw (or say why) and never throw. */
const CASES: Record<string, string> = {
  plain: "flowchart LR\n a --> b --> c",
  td: "graph TD;\n A-->B;\n B-->C;",
  selfLoop: "flowchart LR\n a --> a",
  unknownNodes: "flowchart LR\n a --> ghost\n ghost --> a",
  cycle: "flowchart LR\n a --> b\n b --> c\n c --> a",
  nested: "flowchart LR\n subgraph one[One]\n subgraph two[Two]\n a --> b\n end\n end\n b --> c",
  edgeToGroup: "flowchart LR\n subgraph g[Group]\n a\n end\n a --> g\n g --> z",
  groupAsNode: "flowchart LR\n subgraph g\n a\n end\n g[Group again]\n g --> a",
  sameIdTwice: "flowchart LR\n a[First] --> b\n a[Second] --> c",
  unclosed: "flowchart LR\n subgraph g[Group]\n a --> b",
  extraEnd: "flowchart LR\n a --> b\n end\n end",
  labels: 'flowchart LR\n a["Quoted [x] (y)"] -->|"label | pipe"| b{Decision}\n b -.-> c((Round))',
  icons: "flowchart LR\n a[icon:nope-not-an-icon Name<br/>detail] --> b[fa:fa-database DB]",
  classes: "flowchart LR\n a:::blue --> b:::nonsense\n class a,ghost green",
  styles: "flowchart LR\n a --> b\n classDef blue fill:#f9f\n style a fill:#f9f,stroke:#333\n linkStyle 0 stroke:#f00",
  empty: "flowchart LR",
  noHeader: "a --> b",
  gibberish: "%%{init: {}}%%\n sequenceDiagram\n A->>B: hi",
  unicode: "flowchart LR\n é[Ventes « Europe »] --> 東京[Tokyo]",
  long: "flowchart LR\n" + Array.from({ length: 60 }, (_, i) => ` n${i} --> n${i + 1}`).join("\n"),
  wide: "flowchart LR\n a --> b\n a --> c\n a --> d\n a --> e\n a --> f\n a --> g\n a --> h\n a --> i",
  longLabel: "flowchart LR\n a[" + "very long label ".repeat(30) + "] --> b",
};

describe("diagrams written by a model", () => {
  for (const [name, code] of Object.entries(CASES))
    test(name, () => {
      const graph = parseMermaid(code);
      const svg = renderHtml(createElement(DiagramSvg, { graph }));
      // Drawn, or a message in place of it: never an exception.
      expect(svg.length).toBeGreaterThan(10);
      expect(svg).not.toContain("NaN");
    });
});
