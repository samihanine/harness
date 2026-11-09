# Code structure

Three layers. The **shared code** never imports a feature; a **feature** never imports another
feature; only the **app layer** knows every feature. `bun run typecheck` checks it
(`scripts/check-imports.ts`).

```
src/
  main.tsx, router.tsx    SPA entry (hash history in the single-file export)
  config.ts               every option and flag (Settings › General writes them at runtime; useConfig redraws)
  styles.css              Tailwind v4 tokens, light / dark

  app/                    APP LAYER: the only code that knows every feature
    registry.ts           the features, in the sidebar's order (registerFeatures)
    screen-host.tsx       a feature's screen in its route: loaded on first use, redrawn when an option changes
  routes/                 one file per screen (TanStack Router, file-based): its URL search and its feature
                          (+ __root: sidebar, index: home page, settings: Settings and the features' tabs)

  features/               FEATURES: one folder each, independent of the others
    <name>/module.ts      what the app knows of it (lib/features.ts › Feature): path, label, picture, lazy
                          screen, settings tabs, its collection, what links to Confluence; its hooks
    <name>/screen.tsx     its screen (a Workspace), loaded when opened
    trainer/              report viewer: Explain, Why-changed, guides (+ Settings › Guides)
    builder/              report editing: PBIR tools, editor mode, layout, visual specs, new reports
    data-scientist/       analysis and workbook: pivots, charts, calc sheets, Excel export
    presenter/            real PowerPoint decks: pptx/ (deck.ts: the deck through its XML, template.ts),
                          tools.ts (agent), capture.ts (Power BI pictures), preview.tsx, files.ts (+ Settings › Slides)
    reports/, datasets/   the library screens (datasets: TMDL, Power Query, DAX)
    excel/                Excel tables as databases: table.ts (model), sync.ts (file ⇄ app), xlsx.ts (lists and
                          colors in the file), tools.ts (agent), screen.tsx (Notion-like views)
    confluence/           page agent and screen (markdown editor, attachments)

  lib/                    SHARED: local state and engines
    store.ts              IndexedDB collections (features add theirs: declare module + registerNormalizer)
    features.ts           the registry of features (filled by the app layer)
    library.ts            reports and models of the library, download / import (registerKind, onReportAdded)
    markdown/             syntax.ts (the app's markdown), storage.ts (⇄ Confluence), mermaid.ts, layout.ts (dagre), highlight.tsx
    ooxml/package.ts      Office Open XML files (pptx, xlsx): parts, relationships, content types
    query-spec.ts, mistakes.ts, last.ts, theme.ts, utils.ts
  services/               SHARED: everything that talks to a backend
    ms.ts                 Power BI / Fabric / Graph REST, tokens from the local token server
    pbi.ts                embedding SDKs, DAX (limits, queue, cache)
    excel.ts              Graph workbook API (tables, $batch, look, the app's hidden sheet)
    confluence.ts         Confluence REST API, attachments, the markdown a page was written from
    context-pages.ts      Confluence pages as contexts of the items (localStorage copy)
    report-def.ts, tmdl.ts, tmdl-edit.ts, model.ts, model-samples.ts, report-schema.ts
    hub-config.ts, health.ts
  agent/                  SHARED: the agent harness of every screen
    loop.ts               one conversation per chat, confirmation before changes, JSON tool calls, retries, steps
    reply.ts              tolerant parsing of the model's replies
    llm.ts                the LLM API (never modified)
    context.ts, model-tools.ts, model-editor-tools.ts, pbi-tools.ts, sandbox.ts   shared context and tools
  components/             SHARED UI: workspace, chat, markdown (the one renderer), diagram, badge, mascot, pickers, shadcn ui/
  prompts/                every prompt, one file per feature (imported by the code that sends it)

tests/                    bun test: markdown ⇄ Confluence, mermaid, pptx decks, Excel tables and xlsx (fixtures.ts)
scripts/
  dev.ts                  bun dev: token server + Vite
  token-server.ts         local Microsoft sign-in, GET /get-token; /confluence/<host>/… pass-through
  check-imports.ts        the boundaries above
  screens.ts              every screen's accessibility tree saved / compared (bun run screens [--check])
  export-html.ts          single HTML file
  agent-eval/             agents run outside the browser (offline and live scenarios)
  security/               audit, fixes, library updates
docs/
  performance.md          semantic model limits
  structure.md            this file
```

Adding a screen: a folder in `src/features` with its `module.ts` and `screen.tsx`, its line in
`src/app/registry.ts`, a route file in `src/routes`, its flag in `config.ts` (`SCREENS`).
