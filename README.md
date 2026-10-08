# harness

One TanStack Start app for Power BI work, with an AI assistant (assistant-ui) in the left panel of each screen.

```bash
bun install
bun dev        # http://localhost:3300
```

Sign in once with Microsoft (device code, Microsoft Office public client, no app registration): the same
sign-in gives Power BI, Fabric, Graph and SharePoint tokens. The refresh token stays in `.local/tokens.json`
(server side); the browser only gets a Power BI token for embedding. AI key and model: gear icon.

## Screens

| Screen | What it does |
| --- | --- |
| Library | Reports (a link is enough: pages, visuals, fields and filters are read through embedding, so Viewer / Build rights are enough; clone to My workspace when editable), semantic models (structure read with DAX `INFO.VIEW`), Excel files on SharePoint / OneDrive. Each has a context field given to the AI. Stored in IndexedDB. |
| Viewer | Embedded report. Left: **Info** (links from an Excel table: `label, url, icon, color, description`, icon = lucide name), **Guides** (Excel table: `title, content, links` where links are `pageName` or `pageName/visualName`), **AI** (sees the current page with each visual's data, can change page / filters / slicers, runs DAX). |
| Report builder | Existing report in the Power BI editor (edit mode), or a new one on a semantic model (File › Save as adds it to the library). The AI creates pages and visuals, sets fields, formats, theme, and saves; reports of the library can be given as inspiration. |
| Dataset builder | Excel table in AG Grid (filters, edit, views). Generates a semantic model (TMDL, one table, no relationships) in My workspace with the Excel file as source; "Publish" refreshes it, "Update columns" re-reads the Excel columns and keeps the measures written online. The AI edits rows, filters the grid, reads / writes the model TMDL (types, formats, measures), runs DAX. |

## Code

- `src/server/ms.ts`: sign-in, tokens, one generic Microsoft API call (`ms`), model refresh with fresh SharePoint credentials.
- `src/lib/`: `pbi.ts` (DAX, embedding, what the AI sees of a page), `excel.ts` (Graph workbook tables), `tmdl.ts` (Fabric definitions), `library.ts`, `store.ts` (IndexedDB).
- `src/agent/`: `llm.ts` (deliberately limited plain-text LLM: change the provider at the top), `loop.ts` (JSON tool-call protocol), `pbi-tools.ts`.
- `src/components/chat.tsx`: assistant-ui thread on top of the loop; conversations per screen in IndexedDB.

## Limits

- Model refresh by API: about 8 per day per model with Power BI Pro.
- Visual authoring API: no series colors per visual (use the theme), no text box content.
- Exported visual data needs the "export data" setting allowed on the report / tenant; otherwise the AI uses DAX.
