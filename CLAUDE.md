# Instructions for Claude

Hub v2: a Bun server (`server/`) and a React app (`src/`, TanStack Router + Vite) to work on Power BI
reports, Excel-style analyses linked to Power BI, and PowerPoint decks, with AI assistants that use tools.
Every screen is a `Workspace` (`src/components/workspace.tsx`): the assistant on the left (foldable; its header
chooses the projects), the screen's title and item (renamable), its tabs as pills, its content. Screens: Trainer
(Report / Structure / Semantic model), Builder (Report / Structure / Semantic model / References), Data Scientist
(Analysis / Sources / DAX), Presenter (a deck from the template or an imported PowerPoint: Edit / Present /
Template / Reports / Models), Settings (General / AI / Power BI / Confluence with the projects). Code layout:
`docs/structure.md`.
Run with `bun dev` (app on :3300, API on :3301), check with `bun run typecheck`; `bun run build && bun run start`
serves the built app and the API on one port; `bun run export:server` makes one executable (the app, the
headless browser's scripts and the template's pictures embedded: `scripts/build-server.ts`, `server/static.ts`).

## Architecture

- **The server holds the secrets and calls the services**: Microsoft (Power BI, Fabric, Graph — one
  device-code sign-in with the Microsoft Office public client, no app registration, refresh token in
  `.local/tokens.json`), Confluence, and the AI. The browser only gets a Power BI token for its embeds.
  Only three external dependencies: Microsoft, Confluence, the AI provider. No database, no blob storage:
  settings and files in `.local/` (files cleaned after a few days), everything else in the browser
  (IndexedDB `src/lib/store.ts`, exportable). The server listens on 127.0.0.1 and refuses other origins.
- **LLM layer** (`server/llm/`): one message format (`types.ts`) for three providers, chosen in Settings › AI:
  `maia.ts` (the group's TrustLayer v2, from its documentation), `anthropic.ts` (API key, streaming,
  adaptive thinking, prompt caching, refusal fallback), `claude-code.ts` (Claude Code through the Agent SDK,
  an isolated instance whose only tools are the screen's, given as an in-process MCP server).
- **Agent** (`server/agent/run.ts`): one conversation per chat, append-only (cache and thinking carry over),
  native tool calls, pictures in and out. The first message of a conversation always gets a clarification
  question (`ask_user`, options as buttons) before any change: tools marked `changes` are refused until
  then. State is never pasted in the prompt: the AI reads it with tools (report state, screenshots,
  structures, decks). Screens' prompts and tool sets: `server/agent/screens.ts`, prompts in `src/prompts/`.
- **Tools**: one file per tool in `src/tools/<area>/` (zod args, description, `run` on the server). Tools that
  act on the embedded report call a browser action (`ctx.browser`, implemented in `src/tools/browser/`).
- **Seeing what the user sees** without screen sharing: a headless browser on the server
  (`server/capture/browser.ts`, Playwright with the installed Edge / Chrome) embeds the report with the user's
  view state (a bookmark taken in the browser) and takes a picture. PowerPoint slides are pictured by
  converting the .pptx to PDF through OneDrive (Graph) and drawing the pages with pdf.js.
- **Power BI Pro by default**: no Premium-only API (no Export To File). Model refreshes are never started by
  the AI. Every DAX query goes through `dax()` (`server/pbi/dax.ts`: row limit, timeout, queue, cache) with the
  limits of Settings › Power BI.

## Conventions

- Deterministic code over free generation: the AI fills specs and code writes the result (visuals:
  `server/pbi/visual-spec.ts`, layout: `server/pbi/layout.ts`, decks: markdown → `server/pptx/deck.ts`).
- Builder: the report's PBIR definition is a draft per conversation (`server/pbi/draft.ts`), saved by
  `save_report` (layout checked), then the embed reloads.
- Data Scientist: sources (DAX results, Excel / CSV sheets) become tables of an in-memory SQLite database
  (`server/data/engine.ts`); steps are SQL views; the analysis (`shared/analysis.ts`) is blocks computed from SQL,
  exported to Excel (`server/excel-export.ts`).
- Presenter create: markdown with Slidev-like `---` slides and layouts (`shared/deck.ts`, the template's
  layouts — cover, agenda, section, end — and free ones); the template's pictures in `assets/template/`.
  Presenter edit: the uploaded .pptx changed in place (`server/pptx/ooxml.ts`), each change a new version.
- Context: projects (Settings › Confluence › Projects: a name and a Confluence folder; the default one is
  CONFLUENCE_FOLDER_ID). The projects chosen in the chat header (global, several) give their pages at the start
  of a conversation (`server/projects.ts`, up to a size; the others by title, `read_project_page`).
- Reports and models: their structure is read when they are added and kept in the browser (`src/lib/items.ts`,
  "Refresh" reads it again), sent with the conversation: the tools use it without reading it again. The catalog
  lists workspaces and apps (`server/pbi/catalog.ts`); a pasted link adds an app's report or model.
- Report access (`server/pbi/access.ts`): "edit" (PBIR definition: the Builder's draft tools), "editor" (other
  formats: the Builder works live in the embedded Power BI editor, `src/tools/editor/`, browser side
  `src/tools/browser/editor-actions.ts`), "view". The Builder's "Add" lists editable reports, checked page by page.
- Presenter: Power BI layouts (report, visual, compare, tutorial: `report:`, `page:`, `visual:`, `filters:`) are
  live in full screen (`src/features/presenter/present.tsx`) and pictures in the PowerPoint; an imported
  PowerPoint takes new slides drawn from the template (`Pptx.importSlide`). Template pictures and the vector logo:
  `assets/template/` (cover.jpg, agenda.jpg, section.jpg, end.jpg, logo.svg).
- Settings: the .env file gives the defaults (`.env.example`), changes in Settings are saved over them, "Reset to
  .env" forgets them. Settings › AI and Power BI have "Test the connection".
- UI: shadcn components (`src/components/ui`), tokens in `src/styles.css` (light / dark, primary color from
  Settings), `Workspace`, `Pills`, `Empty`, `Section`. English UI; the AI answers in the user's language.
- Every text sent to the model lives in `src/prompts/` or in a tool's description.
- **Commits**: Sami is the only author. No `Co-Authored-By` line (nor any other mention of Claude) in
  commit messages or pull requests, whatever the default attribution says.
- Write code like the surrounding code: short doc comments saying why, same naming and idioms.
- Verify changes before saying they work: `scripts/checks/agent.ts <screen> '<scope>' '<message>' ['<answer>'…]`
  runs agent turns from the command line (`HUB_PROVIDER` / `HUB_MODEL` to try another provider), `maia-mock.ts`
  checks the TrustLayer format, `flow-*.ts` drive the UI in a headless browser (screenshots in `.local/`),
  `deck.ts` / `deck-render.ts` / `montage.ts` render slides. `bun --watch` restarts the server when a server
  file changes: do not edit server files while a check runs.
