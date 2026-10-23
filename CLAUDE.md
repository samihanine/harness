# Instructions for Claude

TanStack Start app for Power BI work. Every screen is a `Workspace` (`src/components/workspace.tsx`):
the AI assistant on the left (foldable), a header (screen / chosen item, actions) and tabs on the right.
Screens: Viewer (Report / Dataset / Context; the guide of the clicked visual is a tab of the AI panel), Builder (Report / Measures / Reports), Research (Canvas / Report /
Datasets), Reports (Schema / Context), Datasets (Schema / DAX / Context), Excel (Data / Schema /
Context), Contexts, Slides (Preview / Content), Settings (AI / Power BI / General). There is no
library screen: each screen adds and manages its own items (`EntityPicker`, `EntityMenu`). Run with `bun dev` (http://localhost:3300), check with `bun run typecheck`.

## Never modify

- **`src/agent/llm.ts`**: access to the client's private LLM API (provider settings, request format,
  key handling). Never edit it, even to fix or extend something: work around its limits elsewhere
  (`src/agent/loop.ts`, tools, prompts). If a change there seems truly needed, stop and ask. (The
  user once allowed one addition there: file uploads, `uploadAiFile` and `addFileToConversation`.)

## Constraints of every change

- **Generic Microsoft sign-in, any tenant.** One device-code sign-in with the Microsoft Office public
  client (`src/server/ms.ts`): no app registration, no admin consent, no per-tenant setup. Other
  Microsoft resources (Power BI, Fabric, Graph, SharePoint, Forms…) are reached by asking the same
  refresh token for their scope. Never add a solution that needs an app registration, a client secret
  or a tenant-specific setting.
- **Front end only.** All the logic lives in the browser. Server functions (`src/server/`) stay thin
  pass-throughs for Microsoft calls (tokens, CORS, binary files); no business logic, no background
  jobs, no other backend service. The LLM is called from the browser only (no server relay).
- **Limited LLM: build around it.** Plain text in, plain text out: no native tool use, no JSON mode,
  limited context. Structure comes from the harness (`src/agent/loop.ts`: JSON tool-call protocol,
  tools, page state). Keep prompts and contexts short; prefer deterministic code to free-form
  generation: the AI fills a structured spec and code writes the result (new visuals:
  `src/lib/visual-spec.ts`, buttons / conditional formats / reference lines: `src/lib/visual-extras.ts`,
  positions: `src/lib/layout.ts`; DAX queries: `src/lib/query-spec.ts` (`query`, raw DAX stays in
  `run_dax`); Power Query steps: `src/lib/tmdl-edit.ts`; pivots: `src/lib/workbook.ts`).
- **No database, no blob storage.** Files live on SharePoint / OneDrive; data lives in Excel tables
  (`.xlsx`) read and written through the Graph workbook API (`src/lib/excel.ts`). Local state only in
  the browser: IndexedDB (`src/lib/store.ts`: the library, researches, conversations — every library
  entry can be downloaded and imported again) and localStorage (views, last choices). No S3, no SQL,
  no extra storage service. The Researcher's workbooks are files on the user's OneDrive
  (`Hub/Researcher`).
- **Power BI Pro by default.** Do not rely on a Premium / Fabric capacity or on tenant settings that
  are off by default (e.g. the Export To File API). Prefer the client SDKs (`powerbi-client`,
  `powerbi-report-authoring`) and the Fabric definition API (PBIR, TMDL). Model refreshes by API are
  limited (about 8 per day per model): they are started by the user from the UI, never by the AI. When
  a feature depends on something a tenant may not allow, show the service's error clearly instead of
  failing silently or faking a result.
- **Microsoft Forms has no public API.** Forms are read and written through the endpoints the Forms
  designer uses (`readForm` / `formsApi` in `src/server/ms.ts`, `src/lib/forms.ts`): writes only from
  the Schema editor's "Update the form", after the user reviewed the plan. Fragile: show its errors.

## Conventions

- **Everything can be switched off.** Each screen, panel tab and feature button has a flag in
  `src/config.ts` (`SCREENS`, `VIEWER_TABS`, `EXPORT`, `AI_PANEL`, `VISUAL_PREVIEW`…). A new screen, tab
  or feature comes with its flag. Options are read at render time (never copied into module-level
  constants): an admin changes them at runtime from Settings (JSON file on SharePoint, see
  `src/lib/hub-config.ts`). A new option goes into `HubConfig` and the Settings form too. The Settings
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
  that has a model. The planner and the implementer each have their model (Settings › AI). Failed tool
  calls are kept in IndexedDB (`src/lib/mistakes.ts`) and the latest ones of a screen shown to the AI as
  mistakes not to repeat (downloadable from Settings). A big model's context details only the tables the
  request is about (`relevantTables`).
- **Schemas**: a schema is an Excel table (required) + an optional form + the definition of its
  columns, with Forms-compatible types (`src/lib/schema.ts`), and the form's sections and branches.
  It generates both: the form (question title = column name, as Forms names the Excel columns;
  `form: false` keeps a column out of the form) and the Excel file (columns, `_schema` sheet, dropdowns
  and colors), which must keep working as the table receiving the form's answers (new columns only at
  the end). The Schema editor owns the structure; the Excel editor only changes the data. The Graph
  workbook API has no conditional formats and no data validation: those rules are written into the
  file itself (`src/lib/xlsx-rules.ts`, sent back only if the file did not change: eTag). Built-in schemas (guides, links of the report viewer) have fixed
  fields and one Excel file for every report (`report_id` column).
- **Row ids** (`id` column, e.g. `TASK_001`) are given by the app; a user may rename one, always
  checked unique (`idError` in `src/lib/excel.ts`). The AI changes an id only when explicitly asked.
- UI: shadcn components (`src/components/ui`, Tailwind v4 tokens in `src/styles.css`, light/dark via
  `src/lib/theme.ts`), `Workspace` / `SubTabs` / `Section` / `Empty`, choices as badges with a dot
  (`Badge` of `field-input.tsx`). Reuse them rather than new styles.
- **Research** (`src/lib/research.ts`): a canvas (markdown + Power BI visuals embedded from the
  research's reports) and one Power BI report per model whose pages each hold a short text and ONE
  visual, a matrix by default (PBIR written by code from `{chart, rows, columns, values, filters}`), with
  the research's own report measures. Not a second Builder: one visual per page, no formatting.
- **Slides** (`src/slides`): a deck is a short JSON `{title, theme, slides: [{template, ...args}]}`;
  each template (`title.tsx`, `section.tsx`, `table.tsx`…) has a kind (cover, section, subsection,
  summary, content, end: the outline and the content pages' header come from them), zod args, an
  animated component (motion, container units) and a native PowerPoint version (pptxgenjs); the
  header variant and brand letter are in the deck's theme. PDF = print of every slide (Power BI views are
  plain `reportEmbed` iframes so they print).
- **Builder modes**: PBIR (`src/agent/builder.ts`, the definition through Fabric) when the report's
  definition is readable; otherwise "editor mode" (`src/agent/builder-sdk.ts`): the AI edits the
  embedded Power BI editor with `powerbi-report-authoring` (pages, visuals, filters, Top N, frame
  formatting; the inside of visuals through the theme's `visualStyles`), e.g. for reports with an
  encrypted sensitivity label, which Fabric's getDefinition refuses.
- **Report schemas** (`src/lib/report-schema.ts`): kept with each report (`schema`), from the
  definition or, when it is not readable, through a hidden embed; used by the Reports screen and as
  the Builder's reference reports (Reports tab: inspiration context; `copy_report_page` rebuilds one
  of their pages by structure, `src/lib/report-copy.ts`, even when their definition is closed).
- **Builder layout**: visuals are created without positions and placed by `layout_page` (rows); free
  positions only when rows cannot say it. The draft is checked by code before it is saved
  (`draftProblems`: overlaps, out of the page, failing report measures), sent back to the AI once.
- Write code like the surrounding code: short doc comments saying why, same naming and idioms.
- Verify UI changes in the browser preview (dev server on port 3300) before saying they work.
