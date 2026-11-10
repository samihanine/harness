import { describe, expect, test } from "bun:test";
import { fromStorage, toStorage } from "@/lib/markdown/storage";
import { parseMermaid } from "@/lib/markdown/mermaid";

/** Written to Confluence and read back: the same markdown. */
const roundTrip = (md: string) => fromStorage(toStorage(md).storage);

const PAGE = `# Architecture

\`status:green:Live\` \`status:blue:Front end only\`

A single-page app. **All the logic** runs in the *browser* ![](icon:monitor).

> [!NOTE] Rules of every change
> No backend, no database.

## Diagram

\`\`\`mermaid
flowchart LR
  spa[icon:layout-dashboard Hub SPA]:::blue --> token[icon:key-round Token server]:::amber
\`\`\`

\`\`\`ts
const x = await token("powerbi");
\`\`\`

\`\`\`dax
EVALUATE TOPN ( 10, 'Sales' )
\`\`\`

| Layer | Folder |
| --- | --- |
| Screens | \`src/routes\` |
| Features | \`src/features\` |

- one
  - nested
- two

Tasks:

- [ ] to do
- [x] done

> [!WARNING] Refreshes
> About 8 per day.

> [!EXPAND] Details
> Hidden by default.`;

describe("markdown ⇄ Confluence storage", () => {
  test("a page reads back exactly as written", () => {
    expect(roundTrip(PAGE)).toBe(PAGE);
  });

  test("macros and attachments are listed", () => {
    const { storage, files } = toStorage(PAGE);
    expect(storage).toContain('ac:name="status"');
    expect(storage).toContain('ac:name="info"');
    expect(storage).toContain('ac:name="warning"');
    expect(storage).toContain('ac:name="expand"');
    expect(storage).toContain("<ac:task-list>");
    expect(storage).toContain('<ac:parameter ac:name="title">DAX</ac:parameter>');
    expect(files.map((f) => f.kind).sort()).toEqual(["diagram", "icon"]);
    expect(files.find((f) => f.kind === "diagram")!.name).toMatch(/^hub-diagram-[0-9a-f]{8}\.png$/);
  });

  test("a page written in Confluence is read as markdown", () => {
    const html = `<h2>Title</h2><p>Text with <strong>bold</strong> and <a href="https://x.io">a link</a>.</p><ul><li><p>item</p></li></ul><ac:structured-macro ac:name="code"><ac:parameter ac:name="language">python</ac:parameter><ac:plain-text-body><![CDATA[print("a < b")]]></ac:plain-text-body></ac:structured-macro><ac:structured-macro ac:name="tip"><ac:rich-text-body><p>Be quick.</p></ac:rich-text-body></ac:structured-macro><p><ac:emoticon ac:name="tick" ac:emoji-fallback="✅" /> ok</p>`;
    expect(fromStorage(html)).toBe('## Title\n\nText with **bold** and [a link](https://x.io).\n\n- item\n\n```python\nprint("a < b")\n```\n\n> [!TIP]\n> Be quick.\n\n✅ ok');
  });
});

describe("mermaid", () => {
  test("flowchart: groups, labels, icons, colors, chains, details", () => {
    const g = parseMermaid(`flowchart TD
  subgraph front[Front end]
    a[icon:eye Reader<br/>pbi embedded]:::blue
  end
  a -->|DAX| b[(Model)] -.-> c(API)
  c -- token --> d
  d --- a
  class d green`);
    expect(g.direction).toBe("TB");
    expect(g.groups).toEqual([{ id: "front", label: "Front end", parent: undefined }]);
    expect(g.nodes.find((n) => n.id === "a")).toMatchObject({ label: "Reader", icon: "eye", detail: "pbi embedded", color: "blue", parent: "front" });
    expect(g.nodes.find((n) => n.id === "b")).toMatchObject({ label: "Model", icon: "database" });
    expect(g.nodes.find((n) => n.id === "d")?.color).toBe("green");
    expect(g.edges).toEqual([
      { from: "a", to: "b", label: "DAX", dashed: false, line: false },
      { from: "b", to: "c", label: undefined, dashed: true, line: false },
      { from: "c", to: "d", label: "token", dashed: false, line: false },
      { from: "d", to: "a", label: undefined, dashed: false, line: true },
    ]);
  });
});

describe("security of what is rendered and written", () => {
  test("only harmless addresses are kept in links and pictures", async () => {
    const { safeUrl } = await import("@/lib/markdown/syntax");
    expect(safeUrl("https://x.io/a")).toBe("https://x.io/a");
    expect(safeUrl("#part")).toBe("#part");
    expect(safeUrl("javascript:alert(1)")).toBe("");
    expect(safeUrl(" JaVaScRiPt:alert(1)")).toBe("");
    expect(safeUrl("data:text/html;base64,PHNjcmlwdD4=")).toBe("");
    expect(safeUrl("icon:database")).toBe("");
    expect(safeUrl("icon:database", true)).toBe("icon:database");
    expect(safeUrl("data:image/png;base64,iVBORw0KGgo=", true)).toContain("data:image/png");
    expect(toStorage("[x](javascript:alert(1))").storage).not.toContain("javascript:");
  });

  test("a cell starting like a formula is written as text", async () => {
    const { toCell } = await import("@/features/excel/table");
    expect(toCell('=HYPERLINK("http://evil","x")')).toBe(`'=HYPERLINK("http://evil","x")`);
    expect(toCell("@cmd")).toBe("'@cmd");
    expect(toCell("-12.5")).toBe("-12.5");
    expect(toCell("2026-11-12")).toBe("2026-11-12");
    expect(toCell("Plain text")).toBe("Plain text");
  });
});
