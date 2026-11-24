# Code structure

```
server/                 the Bun server (API, AI, Microsoft, Confluence)
  main.ts               routes (/api/…), static app in production, same-origin check
  settings.ts           settings in .local/settings.json (secrets never sent to the browser)
  files.ts              files in .local/files (uploads, captures, PowerPoint versions)
  ms/                   device-code sign-in (auth.ts), Power BI / Fabric / Graph REST (api.ts)
  pbi/                  DAX with limits, model reading, definitions (TMDL, PBIR), report structure,
                        Builder draft, visual specs, layout, measures, catalog of reports
  llm/                  provider-neutral messages; Anthropic, Maia (TrustLayer), Claude Code
  agent/                the agent loop (run.ts), each screen's prompt and tools (screens.ts)
  capture/browser.ts    headless browser: report pictures, embedded structures, PDF pages
  pptx/                 markdown deck → PowerPoint (deck.ts), existing .pptx edits (ooxml.ts), pictures (render.ts)
  data/engine.ts        the Data Scientist's SQL database (SQLite in memory)
  confluence.ts         Confluence REST, markdown ⇄ storage format
  excel-export.ts       an analysis as an Excel workbook
shared/                 types and formats used by both sides (chat events, deck markdown, analyses)
src/
  routes/               one file per screen
  components/           Workspace, ChatPanel, pickers, context links, shadcn ui/
  features/             screen parts: powerbi (embed), views (model, report structure), scientist, presenter, mascot
  tools/                every tool of the AI, one file per tool (run on the server); browser/ = actions on the embed
  prompts/              every prompt
  lib/                  api, store (IndexedDB), chat (SSE), settings, theme
assets/template/        the deck template's pictures (placeholders until the real ones)
scripts/                dev.ts (bun dev), checks/ (headless checks), security/ (audit), make-template-assets.ts
```
