# Code structure

```
src/
  main.tsx, router.tsx    SPA entry (hash history in the single-file export)
  config.ts               every option and flag (Settings › General writes them at runtime)
  styles.css              Tailwind v4 tokens, light / dark
  routes/                 one file per screen (TanStack Router, file-based)
  features/               the code of each feature
    trainer/              report viewer: Explain, Why-changed, guides, visual images
    builder/              report editing: PBIR tools, editor mode, layout, visual specs, drafts
    data-scientist/       canvas, pivot and chart cards
    presenter/            slides-tools.ts (agent) + slides/ (templates, PowerPoint, Power BI embeds)
    datasets/             model editing: TMDL, Power Query steps, data samples
    excel/                Excel data tools, table views
    contexts/             contexts agent
    confluence/           page agent, markdown ⇄ Confluence storage
    mascot/               the hedgehog (SVG layers, outfits, states)
  services/               everything that talks to a backend
    ms.ts                 Power BI / Fabric / Graph REST, tokens from the local token server
    pbi.ts                embedding SDKs, DAX (limits, queue, cache)
    excel.ts              Graph workbook API
    report-def.ts, tmdl.ts, model.ts, report-schema.ts   definitions and structures (Fabric)
    hub-config.ts         options file on SharePoint
    health.ts             sign-in and AI key checks
    confluence.ts         Confluence REST API
  prompts/                every prompt, one file per feature (imported by the code that sends it)
  agent/                  the agent harness shared by every screen
    loop.ts               planner / implementer, JSON tool calls, retries, steps
    reply.ts              tolerant parsing of the model's replies
    llm.ts                the LLM API (never modified)
    context.ts, model-tools.ts, pbi-tools.ts, sandbox.ts   shared context and tools
  lib/                    local state and helpers (store.ts: IndexedDB; library, mistakes, query-spec…)
  components/             shared UI (workspace, chat, pickers, shadcn ui/)
scripts/
  dev.ts                  bun dev: token server + Vite
  token-server.ts         local Microsoft sign-in, GET /get-token
  export-html.ts          single HTML file
  agent-eval/             agents run outside the browser (offline and live scenarios)
  security/               audit, fixes, library updates
docs/
  performance.md          semantic model limits
  structure.md            this file
```
