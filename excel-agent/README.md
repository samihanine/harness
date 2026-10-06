# excel-agent

Edit local Excel files through schemas, with the AI agent (ia-agent) on the side.
Chrome / Edge only (File System Access API).

```bash
bun install
bun dev        # http://localhost:3202 — needs ia-agent running on http://localhost:3200 (VITE_AGENT_URL to change it)
```

- **Schemas** (localStorage): fields with a type (short/long text, number, integer, yes/no, date, option, image),
  options with colors, single or multiple. Import / export as JSON. The agent edits them with JSON Patch.
- **Files**: a local .xlsx linked to a schema (+ optional images folder). Handles are kept in IndexedDB;
  the browser asks again for access after a restart.
- The file is read with ExcelJS; each change touches only its cells, then the workbook is written back
  (debounced). Changes made in Excel are picked up every few seconds.
- Table (inline editing) and form views, search, option filters, sort — remembered per file.
- Agent: sees the schema and the rows of the current view, and adds, updates or deletes rows (every
  change can be undone from the chat). The + button opens a form for a new row.
