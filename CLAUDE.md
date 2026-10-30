# Instructions for Claude

Front-end-only TanStack Router app (Vite SPA) for Power BI work. Every screen is a `Workspace` (`src/components/workspace.tsx`):
the AI assistant on the left (foldable), a header (screen / chosen item, actions) and tabs on the right.
Screens: Viewer "Trainer" (Report / Dataset / Context; the guide of the clicked visual is a tab of the AI panel), Builder (Report / Dataset / Measures / Reports), Research "Data Scientist" (Canvas /
Datasets), Reports (Schema / Context), Datasets (Schema / DAX / Context), Excel (Data / Schema /
Context; hidden by default), Contexts, Confluence (Edit / Preview; hidden by default), Slides "Presenter" (Preview / Edit — slide by slide or the whole deck — / Datasets / Reports), Settings (AI / Power BI / Guides / Slides — the templates — / Performance / Confluence / Mascot / General). Code layout: `docs/structure.md` (features in `src/features/`, backend calls in `src/services/`, every prompt in `src/prompts/`). The home page (`/`, flag `HOME`) shows a card per screen. There is no
library screen: each screen adds and manages its own items (`EntityPicker`, `EntityMenu`). Run with `bun dev` (app on http://localhost:3300 + the token server on :3301), check with `bun run typecheck`. `bun run export:html` (one HTML file), `bun run export:server` (the token server as one executable).

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
  The only other process is the local token server (tokens only, kept in localStorage until they
  expire). No backend, no business logic outside the browser. The LLM is called from the browser only.
- **Limited LLM: build around it.** Plain text in, plain text out: no native tool use, no JSON mode,
  limited context. Structure comes from the harness (`src/agent/loop.ts`: JSON tool-call protocol,
  tools, page state). Keep prompts and contexts short; prefer deterministic code to free-form
  generation: the AI fills a structured spec and code writes the result (new visuals:
  `src/features/builder/visual-spec.ts`, buttons / conditional formats / reference lines: `src/features/builder/visual-extras.ts`,
  positions: `src/features/builder/layout.ts`; DAX queries: `src/lib/query-spec.ts` (`query`, raw DAX stays in
  `run_dax`); Power Query steps: `src/features/datasets/tmdl-edit.ts`).
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

- **Everything can be switched off.** Each screen, panel tab and feature button has a flag in
  `src/config.ts` (`SCREENS`, `VIEWER_TABS`, `EXPORT`, `AI_PANEL`, `VISUAL_PREVIEW`…). A new screen, tab
  or feature comes with its flag. Options are read at render time (never copied into module-level
  constants): an admin changes them at runtime from Settings (JSON file on SharePoint, see
  `src/services/hub-config.ts`). A new option goes into `HubConfig` and the Settings form too. The Settings
  page itself can never be hidden.
- **Agents** (`src/agent/loop.ts`): two phases per request. The planner (read tools only) explores,
  answers at once when nothing has to change, or asks the user with `ask_user` (choices + a plan,
  shown as buttons in the chat) and hands the approved plan to the implementer (read + change tools)
  with `implement`. Agents whose actions are harmless and undone in a click (`direct: true`, e.g. the
  Viewer's filters and pages) skip the planner and act at once. Each screen's agent gives the AI its
  context once (with the library's context files linked to the entries, `src/agent/context.ts`; a
  model comes with its structure and examples of its data or its value profile) and its lists
  (`list_*`: pages, visuals, tables, rows…) as the page state before every model call, so the AI never
  has to ask for them. Changes go to a draft saved when the answer is done (`finish`), with
  `discard_changes` / `undo_last_save`.
  `run_dax` (raw DAX) and `query` (a spec written as DAX by code) are thinking tools on every screen
  that has a model. The planner and the implementer each have their model (Settings › AI), and a maximum of steps (Settings › AI; the planner gets half). Model calls failing for a moment (network, invalid completion, 5xx) are sent again; tool arguments written next to the name are accepted; a refused final tool shows in the trace. Failed tool
  calls are kept in IndexedDB (`src/lib/mistakes.ts`) and the latest ones of a screen shown to the AI as
  mistakes not to repeat (downloadable from Settings). A big model's context details only the tables the
  request is about (`relevantTables`).
- **Excel tables**: the columns are the table's own, read from its header at each read (`readRows`): the
  app never types, adds or moves columns, so the app and the file never differ. The Schema tab only lists
  them and "Sync the formatting" (`formatTable`: grey bold header, borders, widths, wrapped text, rows fitted).
- **Guides** (Trainer): one markdown guide per visual, in IndexedDB (`guides`, `src/features/trainer/guides.ts`),
  downloaded as Excel and imported from Excel / CSV in Settings › Guides. Adding a report creates an empty guide
  for each of its data visuals (`addEmptyGuides`), so the downloaded file lists them all.
- **Row ids** (`id` column, e.g. `TASK_001`) are given by the app; a user may rename one, always
  checked unique (`idError` in `src/services/excel.ts`). The AI changes an id only when explicitly asked.
- **Notices** (a limit of the API, a missing right…): never a banner; `Workspace`'s `notices` show an icon right of the tabs, colored by gravity, opening a drawer with the details.
- UI: shadcn components (`src/components/ui`, Tailwind v4 tokens in `src/styles.css`, light/dark via
  `src/lib/theme.ts`), `Workspace` / `SubTabs` / `Section` / `Empty`, choices as badges with a dot
  (`Badge` of `field-input.tsx`). Reuse them rather than new styles.
- **Data Scientist** (`src/features/data-scientist/research.ts`): a canvas — a markdown text (the answer, written there by
  default) and cards below it, ONE table or chart each on one model: a pivot (rows, columns, values,
  filters; DAX written by code) or a chart (recharts, from free DAX). No Power BI report.
- **Slides** (`src/features/presenter/slides`): a deck is a short JSON `{title, theme, slides: [{template, ...args}]}`;
  each template (`title.tsx`, `agenda.tsx`, `table.tsx`…) has a kind (cover, section — an agenda naming its part —, subsection,
  summary, content, end: the outline and the content pages' header come from them), zod args, an
  animated component (motion, container units) and a native PowerPoint version (pptxgenjs); the
  header variant and brand letter are in the deck's theme. Images: public links or inline base64 in
  the JSON. Power BI: each report page embedded once (SDK, app token) in a layer over the slide shown
  and reused by every slide of that page (`src/features/presenter/slides/powerbi.tsx`); before a PDF / PowerPoint export
  the tab is captured (`src/features/presenter/slides/capture.ts`) so the views go in as pictures. PowerPoint: text and
  simple shapes stay editable, the rest as pictures; Morph transition on every slide. Diagrams
  (`src/features/presenter/slides/diagram.ts`): boxes, groups and arrows laid out by dagre, drawn by the template.
- **Prompts**: every text sent to the model lives in `src/prompts/<feature>.ts` (constants, or small functions
  when they depend on the screen), imported by the code that sends it.
- **Confluence** (`src/services/confluence.ts`, `src/features/confluence/`): the pages of one folder (Settings ›
  Confluence: route URL, key, folder id; kept in the browser), edited as markdown (converted from / to the
  storage format), written as the page's next version. Called from the browser: the route must allow CORS.
- **Builder modes**: PBIR (`src/features/builder/builder.ts`, the definition through Fabric) when the report's
  definition is readable; otherwise "editor mode" (`src/features/builder/builder-sdk.ts`): the AI edits the
  embedded Power BI editor with `powerbi-report-authoring` (pages, visuals, filters, Top N, frame
  formatting; the inside of visuals through the theme's `visualStyles`), e.g. for reports with an
  encrypted sensitivity label, which Fabric's getDefinition refuses.
- **Report schemas** (`src/services/report-schema.ts`): kept with each report (`schema`), from the
  definition or, when it is not readable, through a hidden embed; used by the Reports screen and as
  the Builder's reference reports (Reports tab: inspiration context; `copy_report_page` rebuilds one
  of their pages by structure, `src/features/builder/report-copy.ts`, even when their definition is closed).
- **Builder layout**: visuals are created without positions and placed by `layout_page` (rows); free
  positions only when rows cannot say it. The draft is checked by code before it is saved
  (`draftProblems`: overlaps, out of the page, failing report measures), sent back to the AI once.
- **Security** (`scripts/security`): `bun run audit` (known vulnerabilities of the dependencies + secrets
  committed by mistake), `bun run audit:fix` (smallest safe versions, overrides for nested copies),
  `bun run update:libs` (`--dry` to list): dependency changes are kept only if typecheck + build pass.
- **Commits**: Sami is the only author. No `Co-Authored-By` line (nor any other mention of Claude) in
  commit messages or pull requests, whatever the default attribution says.
- **Mascot** (`src/features/mascot`): a hedgehog in SVG layers (React + motion): outfits fill slots of the body
  (`outfits.tsx`: diver, trainer, builder, datascientist, presenter), states move its pivots (`mascot.tsx`:
  idle, thinking, talking, working, sleeping), so any outfit plays any state. The chat shows it left of the
  last answer, dressed for the screen (`outfitOf`) and in the agent's state, asleep when nothing happens
  (awake under the mouse: `WakingMascot`); the sidebar and the home page show it for the assistants
  (`screen-logo.tsx`); Settings › Mascot previews it.
- Write code like the surrounding code: short doc comments saying why, same naming and idioms.
- Verify UI changes in the browser preview (dev server on port 3300) before saying they work.
