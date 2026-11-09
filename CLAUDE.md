# Instructions for Claude

Front-end-only TanStack Router app (Vite SPA) for Power BI work. Every screen is a `Workspace` (`src/components/workspace.tsx`):
the AI assistant on the left (foldable), a header (screen / chosen item, actions) and tabs on the right.
Screens: Viewer "Trainer" (Report / Dataset / Context; the guide of the clicked visual is a tab of the AI panel), Builder (Report / Dataset / Measures / Reports), Research "Data Scientist" (Canvas: the analysis and a workbook /
Datasets), Slides "Presenter" (Slides / Outline / Datasets / Reports: a real .pptx), Reports (Schema / Context), Datasets (Schema / DAX / Context), Excels (Data — Notion-like views — / Columns /
Context), Confluence (Edit — markdown and preview side by side — / Preview / Linked to), Settings (AI / Power BI / the features' tabs: Guides, Slides — the template deck — / Performance / Confluence / Mascot / General). Code layout: `docs/structure.md` (one folder per feature in `src/features/`, each described by its `module.ts`; shared code in `src/lib`, `src/services`, `src/agent`, `src/components`; every prompt in `src/prompts/`; the app layer `src/app` + `src/routes` is the only one that knows every feature). The home page (`/`, flag `HOME`) shows a card per screen. There is no
library screen: each screen adds and manages its own items (`EntityPicker`, `EntityMenu`). Run with `bun dev` (app on http://localhost:3300 + the token server on :3301), check with `bun run typecheck` (types + the boundaries between modules, `scripts/check-imports.ts`) and `bun test` (converters, decks, Excel tables); `bun run screens` / `bun run screens --check` (dev server running) saves / compares each screen's accessibility tree, for changes that must not change what users see. `bun run export:html` (one HTML file), `bun run export:server` (the token server as one executable).

## Never modify

- **`src/agent/llm.ts`**: access to the client's private LLM API (provider settings, request format,
  key handling). Never edit it, even to fix or extend something: work around its limits elsewhere
  (`src/agent/loop.ts`, tools, prompts). If a change there seems truly needed, stop and ask. (The
  user once allowed one addition there: file uploads, `uploadAiFile` and `addFileToConversation`.)

## Constraints of every change

- **Generic Microsoft sign-in, any tenant.** One device-code sign-in with the Microsoft Office public
  client, done by the local token server (`scripts/token-server.ts`, `GET /get-token?scope=…`): no app registration, no admin consent, no per-tenant setup. Other
  Microsoft resources (Power BI, Fabric, Graph, SharePoint, Forms…) are reached by asking the same
  refresh token for their scope. Never add a solution that needs an app registration, a client secret
  or a tenant-specific setting.
- **Front end only.** All the logic lives in the browser, which calls the Microsoft APIs directly
  (`src/services/ms.ts`: REST of Power BI / Fabric / Graph; `src/services/pbi.ts`: the embedding SDKs and DAX).
  The only other process is the local token server (tokens, kept in localStorage until they
  expire, and a pass-through to Confluence Cloud, which allows no browser calls). No backend, no business logic outside the browser. The LLM is called from the browser only.
- **Limited LLM: build around it.** Plain text in, plain text out: no native tool use, no JSON mode,
  limited context. Structure comes from the harness (`src/agent/loop.ts`: JSON tool-call protocol,
  tools, page state). Keep prompts and contexts short; prefer deterministic code to free-form
  generation: the AI fills a structured spec and code writes the result (new visuals:
  `src/features/builder/visual-spec.ts`, buttons / conditional formats / reference lines: `src/features/builder/visual-extras.ts`,
  positions: `src/features/builder/layout.ts`; DAX queries: `src/lib/query-spec.ts` (`query`, raw DAX stays in
  `run_dax`); Power Query steps: `src/services/tmdl-edit.ts`).
- **No database, no blob storage.** Files live on SharePoint / OneDrive; data lives in Excel tables
  (`.xlsx`) read and written through the Graph workbook API (`src/services/excel.ts`). Local state only in
  the browser: IndexedDB (`src/lib/store.ts`: the library, researches, conversations — every library
  entry can be downloaded and imported again) and localStorage (views, last choices). No S3, no SQL,
  no extra storage service.
- **Power BI Pro by default.** Do not rely on a Premium / Fabric capacity or on tenant settings that
  are off by default (e.g. the Export To File API). Prefer the client SDKs (`powerbi-client`,
  `powerbi-report-authoring`) and the Fabric definition API (PBIR, TMDL). Model refreshes by API are
  limited (about 8 per day per model): they are started by the user from the UI, never by the AI. When
  a feature depends on something a tenant may not allow, show the service's error clearly instead of
  failing silently or faking a result.
- **Semantic models' performance first** (`docs/performance.md`): every DAX query goes through `dax()`
  (`src/services/pbi.ts`: row limit by TOPN, abort after a few seconds, queue per model, short cache), with
  the limits of Settings › Performance (`PERFORMANCE`); the AI is told them (`daxLimits`).

## Conventions

- **Modules** (`docs/structure.md`): a feature never imports another feature, the shared code never imports a
  feature (`bun run typecheck` refuses it). A feature describes itself in `module.ts` (`Feature` of
  `src/lib/features.ts`: path, label, picture, its lazy screen — its own chunk —, settings tabs, the collection its
  conversations are about, what can be linked to Confluence pages) and plugs into the shared code through
  registries: its library collection (`declare module "@/lib/store" { interface Collections … }`,
  `registerNormalizer`), downloadable kinds (`registerKind`), hooks (`onReportAdded`). `src/app/registry.ts` lists the
  features; each route file hosts one (`ScreenHost`). A change of option redraws the screens in place
  (`useConfig`), never mounts them again.
- **Everything can be switched off.** Each screen, panel tab and feature button has a flag in
  `src/config.ts` (`SCREENS`, `VIEWER_TABS`, `EXPORT`, `AI_PANEL`, `VISUAL_PREVIEW`…). A new screen, tab
  or feature comes with its flag. Options are read at render time (never copied into module-level
  constants): an admin changes them at runtime from Settings (JSON file on SharePoint, see
  `src/services/hub-config.ts`). A new option goes into `HubConfig` and the Settings form too. The Settings
  page itself can never be hidden.
- **Agents** (`src/agent/loop.ts`): one model conversation per chat, with every tool of the screen (the
  context is sent once, then only what changed; it is read ahead when the chat opens, so the first
  message does not wait for model samples or Confluence). Always EXPLORE (read tools, DAX: names, values,
  figures), then CLARIFY or CONFIRM (`ask_user`: what was found, questions with the real options, shown
  as buttons) — never "not possible" on the screen —, then ACT on the answer. On a screen with a model an
  `ask_user` before any DAX is sent back once; tools that change things are refused until confirmed,
  unless the request is obvious (the model says so with `proceed` in the same reply). Agents whose
  actions are harmless and undone in a click (`direct: true`, e.g. the Viewer's filters and pages) or
  expected without asking (`autoImplement`, the Data Scientist's canvas) act at once. Each screen's
  agent gives the AI its context once (with the Confluence pages linked to the entries,
  `src/agent/context.ts`; a model comes with its structure and examples of its data or its value
  profile) and its lists (`list_*`: pages, visuals, tables, rows…) as the page state before every
  model call, so the AI never has to ask for them. Changes go to a draft saved when the answer is done
  (`finish`), with `discard_changes` / `undo_last_save`; `beforeAnswer` sends back what is still
  missing (e.g. the Data Scientist's pivot).
  `run_dax` (raw DAX) and `query` (a spec written as DAX by code) are thinking tools on every screen
  that has a model. One model and a maximum of steps per message (Settings › AI); a long conversation
  is restarted (with the chat as text) so calls stay fast. Model calls failing for a moment (network, invalid completion, 5xx) are sent again; tool arguments written next to the name are accepted; a refused final tool shows in the trace. Failed tool
  calls are kept in IndexedDB (`src/lib/mistakes.ts`) and the latest ones of a screen shown to the AI as
  mistakes not to repeat (downloadable from Settings): those naming data only in the same scope (another
  model's names would mislead), a "worked instead" only when the fix is a call about the same thing. A
  reply repeating its "tools" key keeps every call (`reply.ts`). DAX tools: `op` aliases ("=", "!="),
  `dataset_id` of the only model accepted, and a query returning no rows says when a filter value does
  not exist, with the close values that do (`missingValues`). A big model's context details only the tables the
  request is about (`relevantTables`).
- **Excel tables** (`src/features/excel`): a table on SharePoint used as a simple database, Notion-like (views as
  tabs — table, board, list —, filter chips, a row opened on the side). 4 column types only (`table.ts`: text, long
  text, date yyyy-mm-dd, option with a color). Rows are written at once (Graph `$batch`), each found again in a
  fresh read of the file (never overwritten by position); the AI changes rows BY CONDITION (`where`/`set`, no ids).
  The structure is written by one sync pass after each column change (`sync.ts`): columns added / deleted /
  reordered / renamed, then each option column's drop-down list and colors written into the .xlsx itself
  (`xlsx.ts`, data validation + conditional formatting: the Graph API has neither), the look (`formatTable`: grey
  bold header, borders, widths by type, wrapped text, rows fitted, header frozen), and the types and views kept in
  a hidden sheet `_hub` of the file. The file's version is watched (20 s): a change made in Excel is read again.
- **Guides** (Trainer): one markdown guide per visual, in IndexedDB (`guides`, `src/features/trainer/guides.ts`),
  downloaded as Excel and imported from Excel / CSV in Settings › Guides. Adding a report creates an empty guide
  for each of its data visuals (`addEmptyGuides`), so the downloaded file lists them all.
- **Notices** (a limit of the API, a missing right…): never a banner; `Workspace`'s `notices` show an icon right of the tabs, colored by gravity, opening a drawer with the details.
- UI: shadcn components (`src/components/ui`, Tailwind v4 tokens in `src/styles.css`, light/dark via
  `src/lib/theme.ts`), `Workspace` / `SubTabs` / `Section` / `Empty`, choices as badges with a dot
  (`Badge` of `components/badge.tsx`). Reuse them rather than new styles.
- **Data Scientist** (`src/features/data-scientist/research.ts`): the analysis (markdown) on top and a workbook below,
  shown as a simplified Excel window with its sheets as tabs at the bottom: pivot (rows, columns, values, filters;
  DAX written by code), chart (a pivot's parameters drawn, recharts) and calc (cells with Excel formulas, computed
  in the app by `formulas.ts`; pivots read with GETPIVOTDATA, so a figure the models lack is rebuilt in the open).
  Every request: explored with DAX, confirmed by the user, then at least one pivot answering it (never one for
  show). Older canvases (cards) become sheets. The workbook downloads as Excel (`pivot-excel.ts`, written by hand):
  a pivot of model measures filtered on lists of values is a PivotTable ON the model (OLAP, as "Analyze in Excel");
  any other, a "… data" table connected to its model (OLE DB, running the sheet's DAX) and a PivotTable on it; calc
  sheets keep their formulas (GETPIVOTDATA rewritten for the PivotTables — model names on OLAP ones —, recomputed on
  open). A pivot's fields are edited on the right of its sheet (Filters / Columns / Rows / Values).
- **Presenter** (`src/features/presenter`): any real PowerPoint (no template required), from this computer
  (saved back in place where the browser allows it, else downloaded) or SharePoint (read and written through
  Graph, only if nobody saved it meanwhile). The deck is changed through its XML (`pptx/deck.ts` on
  `src/lib/ooxml`): the AI never sees the slides, it reads them as text — each shape with its id, kind, place in
  %, size and room for text — and uses small tools (`set_text` keeps the shape's look, `replace_text`,
  `add_slide` from a type of the template deck, `duplicate_slide`, `insert_visual`, `set_box`…), plus
  `read_xml` / `write_xml` on any part for anything else. The template deck (Settings › Slides) is a .pptx where
  each slide is a type (named in its notes: "type: Two columns"); its slides are copied with their pictures,
  taking this deck's layout of the same name. Power BI visuals go in as pictures: embedded over the screen
  for a moment and cut from a capture of the tab (`capture.ts`, shared once per session; works on Pro). Each
  change is kept in IndexedDB with the earlier versions (undo); the slides are drawn from their XML
  (`preview.tsx`).
- **Prompts**: every text sent to the model lives in `src/prompts/<feature>.ts` (constants, or small functions
  when they depend on the screen), imported by the code that sends it.
- **Markdown** (`src/lib/markdown`, one renderer: `src/components/markdown.tsx`, for answers, guides, analyses and
  pages): GitHub markdown plus constructs Confluence can hold (`syntax.ts`): ```mermaid diagrams (a subset of
  flowcharts parsed by `mermaid.ts`, laid out by dagre in `layout.ts`, drawn as SVG with lucide icons by
  `components/diagram.tsx`), code colored by language (`highlight.tsx`), `> [!NOTE]` / TIP / WARNING / IMPORTANT /
  EXPAND panels, `` `status:green:Done` ``, `![](icon:name)`, task lists. `storage.ts` converts to and from the
  Confluence storage format without loss: each construct has a macro whose parameters keep it (a diagram: its
  picture attached + its source in an expand macro; DAX / M / TMDL: the code macro's title), round trip tested.
- **Confluence** (`src/services/confluence.ts`, `src/features/confluence/`): the pages of one folder (Settings ›
  Confluence: route URL, key, folder id; kept in the browser), edited as markdown, written as the page's next
  version with its diagrams and icons attached as PNG (only the missing ones: their names are their hashes). A
  page not changed in Confluence since reads back as the exact markdown it was written from (`markdownOf`). Each
  page is a context: an item's Context tab (or the page's "Linked to" tab, filled from the features' `linkable`)
  links pages (`contextIds`, kept in the browser); their markdown is kept in localStorage
  (`src/services/context-pages.ts`), read again once per conversation, the copy used when Confluence
  fails. Called from the browser: the route must allow CORS;
  Confluence Cloud does not, so the token server forwards to it (`/confluence/<site>.atlassian.net/…`, Atlassian
  hosts only, the user's key passed through).
- **Builder modes**: PBIR (`src/features/builder/builder.ts`, the definition through Fabric) when the report's
  definition is readable; otherwise "editor mode" (`src/features/builder/builder-sdk.ts`): the AI edits the
  embedded Power BI editor with `powerbi-report-authoring` (pages, visuals, filters, Top N, frame
  formatting; the inside of visuals through the theme's `visualStyles`), e.g. for reports with an
  encrypted sensitivity label, which Fabric's getDefinition refuses.
- **Report schemas** (`src/services/report-schema.ts`): kept with each report (`schema`), from the
  definition or, when it is not readable, through a hidden embed; used by the Reports screen and as
  the Builder's reference reports (Reports tab: inspiration context; `copy_report_page` rebuilds one
  of their pages by structure, `src/features/builder/report-copy.ts`, even when their definition is closed).
- **Builder edits**: `set_visual_fields` also changes a visual's type in place (`type`), read back from the editor.
- **Builder formatting**: `style_visual` (`src/features/builder/visual-style.ts`): title, field display names,
  series colors, text size and color, labels, legend, header colors, frame — a plain spec written by code (its
  keys also accepted next to "visuals"); `format_visual` for any other property, `set_conditional_format` for
  colors by value. A role named as another type's ("Values" on a cardVisual) is mapped (`fitRoles`); a visual
  whose fields are refused is removed, never left half made.
- **Builder layout**: visuals are created without positions and placed by `layout_page` (rows); free
  positions only when rows cannot say it. The draft is checked by code before it is saved
  (`draftProblems`: overlaps, out of the page, failing report measures), sent back to the AI once.
- **Security** (`scripts/security`): `bun run audit` (known vulnerabilities of the dependencies + secrets
  committed by mistake), `bun run audit:fix` (smallest safe versions, overrides for nested copies),
  `bun run update:libs` (`--dry` to list): dependency changes are kept only if typecheck + build pass.
- **Commits**: Sami is the only author. No `Co-Authored-By` line (nor any other mention of Claude) in
  commit messages or pull requests, whatever the default attribution says.
- **Mascot** (`src/components/mascot`): a hedgehog in SVG layers (React + motion): outfits fill slots of the body
  (`outfits.tsx`: diver, trainer, builder, datascientist, presenter), states move its pivots (`mascot.tsx`:
  idle, thinking, talking, working, sleeping), so any outfit plays any state. The chat shows it left of the
  last answer, dressed for the screen (`outfitOf`) and in the agent's state, asleep when nothing happens
  (awake under the mouse: `WakingMascot`); the sidebar and the home page show it for the assistants
  (`screen-logo.tsx`); Settings › Mascot previews it.
- Write code like the surrounding code: short doc comments saying why, same naming and idioms.
- Verify UI changes in the browser preview (dev server on port 3300) before saying they work.
