# excel-agent

Edit Excel files through schemas, with the AI agent (ia-agent) on the side. Files and image folders
are on this computer (Chrome / Edge) or on SharePoint / OneDrive, chosen per file.

```bash
bun install
bun dev        # http://localhost:3202 + local API (Microsoft sign-in) — ia-agent expected on :3200 (VITE_AGENT_URL)
bun start      # production build served with the API by Bun
bun test tests # storage tests (SharePoint backend against a simulated Graph API)
```

## Storage (`src/storage`)

The app only uses two interfaces: `TableBackend` (rows of a sheet, keyed by `id`) and `ImageStore`.

| | This computer | SharePoint / OneDrive |
| --- | --- | --- |
| Excel | File System Access + ExcelJS: changes touch their cells, then the file is written (an .xlsx is a zip) | Graph Excel API in place: one persistent workbook session, rows kept as an Excel table, each save = one `$batch` round trip (row writes, appends, deletes) |
| Images | files in a folder, cells hold the file name | uploaded to a folder, cells hold a sharing link (anonymous when allowed, else organization) |
| Outside changes | file date, checked every 3 s | table values, checked every 15 s |
| Rich text | native rich runs | plain text (Graph cannot format characters) |

Sign-in: device code once (Microsoft Office public client, any tenant); the local API (`server/`) keeps
the refresh token in `.local/tokens.json` and renews access tokens. Image cells may also hold any web link.

- **Schemas** (localStorage): fields with a type (short/long text, number, integer, yes/no, date, option, image),
  options with colors, single or multiple. Import / export as JSON. The agent edits them with JSON Patch.
- **Files**: a local .xlsx linked to a schema (+ optional images folder). Handles are kept in IndexedDB;
  the browser asks again for access after a restart.
- The file is read with ExcelJS; each change touches only its cells, then the workbook is written back
  (debounced). Changes made in Excel are picked up every few seconds.
- Table (inline editing) and form views, search, option filters, sort — remembered per file.
- Agent: sees the schema and the rows of the current view, and adds, updates or deletes rows (every
  change can be undone from the chat). The + button opens a form for a new row.
