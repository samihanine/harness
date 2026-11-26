# Code structure

```
server/                 the Bun server (API, AI, Microsoft, Confluence)
  main.ts               routes (/api/…), static app in production, same-origin check
  settings.ts           settings in .local/settings.json (secrets never sent to the browser)
  files.ts              files in .local/files (uploads, captures, PowerPoint versions)
  ms/                   device-code sign-in (auth.ts), Power BI / Fabric / Graph REST (api.ts)
  pbi/                  DAX with limits, model reading, definitions (TMDL, PBIR, PBIR-Legacy as PBIR: legacy.ts),
                        report structure, Builder draft and .pbix (pbix.ts), visual specs, layout, catalog, resolve
  llm/                  provider-neutral messages; Anthropic, Maia (TrustLayer), Claude Code
  agent/                the agent loop (run.ts), each screen's prompt and tools (screens.ts)
  capture/browser.ts    headless browser: report pictures, embedded structures, PDF pages
  slides/render.ts      presentations: pictures, PDF and native PowerPoint of templates × markdown
  pptx/                 reading an attached .pptx (ooxml.ts) and its pictures through OneDrive (render.ts)
  workbook/             the Data Scientist's workbook: Excel's API, its file's XML, its HTML preview
  teams/                the Teams assistant: Graph chat calls (api.ts), the watching loop (bot.ts)
  library.ts            the catalog of reports and models (Confluence pages)
  confluence.ts         Confluence REST, markdown ⇄ storage format
  projects.ts           the projects' Confluence pages (kept in .local/projects) as a conversation's context
shared/                 types and formats used by both sides (chat events, slides and templates, links)
src/
  routes/               one file per screen
  components/           Workspace, ChatPanel, pickers, screen icons, shadcn ui/
  features/             screen parts: powerbi (embed), views (model, report structure), scientist, presenter, mascot
  tools/                every tool of the AI, one file per tool (run on the server); editor/ = the Builder's editor
                        mode; browser/ = actions on the embedded report (view and editor)
  prompts/              every prompt
  lib/                  api, store (IndexedDB), items (structures kept), chat (SSE, runs per conversation),
                        params (item and conversation in the address), projects, settings, theme
assets/template/        the deck template's pictures (placeholders until the real ones)
scripts/                dev.ts (bun dev), checks/ (headless checks), security/ (audit), make-template-assets.ts
```
