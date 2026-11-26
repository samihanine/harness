# Instructions for Claude

Hub v2: a Bun server (`server/`) and a React app (`src/`, TanStack Router + Vite) to work on Power BI
reports, Excel-style analyses linked to Power BI, and PowerPoint decks, with AI assistants that use tools.
Every screen is a `Workspace` (`src/components/workspace.tsx`): the assistant on the left (foldable, resizable; its
header chooses the projects, the conversation, its feedback), the screen's title and item (renamable), its tabs as
pills, its content. Screens: Trainer (Report / Structure / Semantic model), Builder (a dropped .pbix: Report /
Structure / Semantic model / References), Data Scientist (Workbook / Semantic models / Excel files / DAX), Presenter
(Content / Visual / Present / Reports / Models), Teams (the assistant of one Teams chat), Settings (General / AI /
Power BI / Confluence, and Reports / Models when the catalog is editable). Code layout: `docs/structure.md`.
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
  `server/pbi/visual-spec.ts`, layout: `server/pbi/layout.ts`, slides: templates × markdown → `server/slides/render.ts`).
- Builder: the user drops a .pbix (live-connected to a model: the Build permission is enough); it is imported privately
  into their My workspace (Imports API, `server/pbi/pbix.ts`) and embedded; its definition is a draft per conversation
  (`server/pbi/draft.ts`) saved by `save_report` (layout of the changed pages checked); the .pbix downloaded is the
  dropped file with the copy's definition put back (security bindings removed). The old format (PBIR-Legacy) is shown
  to the tools as PBIR and written back, only what changed (`server/pbi/legacy.ts`, `scripts/checks/legacy.ts`).
  Designed pages: text boxes, pictures by link, shapes, buttons with icons and actions, free positions with layers
  (`place_visuals`), one visual's look, series colors, colors from a measure (`format_visual`), report measures
  (`report_measure`): the model is never changed. Checks: `scripts/checks/pbix.ts`, `flow-pbix.ts`.
- Trainer: a hierarchy slicer takes a tree (a date range becomes its years, months, days); a page filter is set only
  once the refusing slicer is cleared. `visual_data` falls back to DAX from the visual's fields and filters when
  Power BI refuses the export. Figures are checked with DAX. The report fits the width of its area (its "100 %").
- Data Scientist: an Excel workbook in the user's OneDrive ("Hub workbooks"), changed through Excel's API (Graph
  workbook: `server/workbook/excel.ts`) and computed by Excel itself — CUBE formulas on the Power BI models (connections
  put in the file), formulas over big uploaded sheets; conditional formats and drop-down lists written in the file's XML
  (`server/workbook/ooxml.ts`, Graph v1.0 has none); shown as HTML from Excel's texts and the file's formats
  (`server/workbook/preview.ts`); downloaded working. Tools: `src/tools/workbook/`. Check: `scripts/checks/workbook.ts`.
- Presenter: the look apart from the content (`shared/slides.ts`). Templates: elements in inches bound to slots (text,
  shape, image, powerbi, repeated items), kept in the browser (favorites, export / import; the company's in
  `shared/slide-templates.ts`, pictures in `assets/template/`). Content: markdown, one slide per `---` block
  (`template:`, `# Title`, body, `![](…)`, Power BI links of a page or a visual, `**value** label` items). The same HTML
  draws the app, the AI's pictures and the PDF (headless browser: no Microsoft call); the PowerPoint is native and
  editable (pptxgenjs, `server/slides/render.ts`). Power BI views are live in full screen. An attached PowerPoint is read
  (`read_powerpoint`) to recreate its look as templates. Check: `scripts/checks/slides.ts`.
- Teams (`server/teams/`): one chat (group or one-to-one) of the signed-in user, its last message checked every 4 s;
  new messages go to an agent that answers or says NO_REPLY (`chatMember`: no clarification question), reading only;
  its answers are signed, short, with pivot-like data tables (`data_table`) and the mascot's GIFs (made in the app,
  `.local/gifs`). Its own messages are never answered (their ids kept in `.local/teams.json`). Check:
  `scripts/checks/teams-mock.ts` (a fake chat).
- Context: projects are the sub-folders of the Confluence projects folder (CONFLUENCE_PROJECTS_FOLDER_ID); the ones
  chosen in the chat header give their pages whole at the start of a conversation (`server/projects.ts`, up to 150 000
  characters, the rest by title with `read_project_page`); read once, kept in `.local/projects`, "Refresh" reads again.
- Catalog: the team's reports and models are Confluence pages (CONFLUENCE_REPORT_FOLDER_ID, CONFLUENCE_DATASET_FOLDER_ID:
  title = name, content = Power BI link + description; `server/library.ts`). The screens list only them (hidden ones per
  user); they are changed in Settings › Reports / Models when CONFLUENCE_CATALOG_EDIT is on. An item is opened (its
  rights checked) when first used; its structure is read then and kept in the browser (`src/lib/library.ts`,
  `src/lib/items.ts`). Any Power BI link works (an app's report finds its model by `datasetWorkspaceId`). Check:
  `scripts/checks/confluence-mock.ts` with `flow-round5.ts`, `projects-mock.ts`.
- Conversations: the item and the conversation are in the address (`?item=…&c=…`, `src/lib/params.ts`), so each
  browser tab has its own; each conversation runs on its own (several at once). The history lists every
  conversation of the assistant: opening one shows its item and the report's view kept with it (a bookmark).
  Conversations can be rated, commented and downloaded as JSON.
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
  `slides.ts` / `workbook.ts` / `pbix.ts` / `teams-mock.ts` check each screen's engine. `bun --watch` restarts the server when a server
  file changes: do not edit server files while a check runs.
