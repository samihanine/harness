# pbi-agent

Power BI workbench with the AI agent (ia-agent) in the left panel. Chrome / Edge (local files).

```bash
bun install
bun dev      # http://localhost:3201 — front-end + local API in one process (ia-agent expected on :3200)
bun start    # production build served with the API by Bun
```

## Pages

| Page | What it does |
| --- | --- |
| Settings | Power BI sign-in (device code, any tenant, Microsoft "Power BI" public client). The local API keeps the refresh token in `.local/tokens.json` and renews access tokens itself. Database folder. |
| Database | `reports.xlsx`, `datasets.xlsx`, `guides.xlsx` in a local folder. Rename, context, refresh, delete. Add a report from a **.pbix** (full definition) or a **link** (definition downloaded when allowed, else read through embedding); its dataset is added automatically. Add a local dataset from an **Excel** file (sheets → tables, converted to Excel tables when needed). Large JSON lives in `json/<table>/<id>.<kind>.json` next to the Excel files (a cell holds at most 32,767 characters). |
| DAX | Query runner on any dataset: DAX for Power BI models, SQL for local Excel models. Results downloadable (xlsx / csv). |
| Research | Several sources at once, pivot table (rows / columns / values / filters) built in the UI or by the agent; rows are mapped to a column of each source, so sources can be compared; new measures (native, or across sources: `{A} - {B}`) can be switched off. |
| Builder | One visual per row: live preview (DAX + ECharts) on the left, settings on the right (type, fields by role, filters, format, position). Downloads a .pbix connected live to the dataset; "Open in Power BI" imports a copy in My workspace and shows the real visuals. |
| Viewer | Embedded report (your own token). Guide tab: markdown guides linked to one or more visuals; a clicked visual opens its guide or offers to attach / create one. |

## Model reading with partial rights

Everything goes through `executeQueries` (Build / Read permission is enough): `INFO.VIEW.*` for tables, columns,
measures and relationships; `COLUMNSTATISTICS()` when columns are not listable; measures seen in the reports
when their definitions are not readable. What fails is listed in the model's `warnings` instead of failing.

## Server

`server/app.ts` (Hono): `/api/status`, `/api/login`, `/api/token`, `/api/logout`, and `/api/pbi/*` / `/api/fabric/*`
pass-throughs with the token (for the endpoints browsers cannot call: report export, imports). Local origin only.
