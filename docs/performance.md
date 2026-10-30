# Performance of the semantic models

The assistants query the user's semantic models with DAX (Power BI REST API `executeQueries`). These
models share their capacity with every report user, so the app follows Microsoft's guidance for
query load. The limits are set in **Settings › Performance** (`PERFORMANCE` in `src/config.ts`).

| Limit | Default | What it does |
| --- | --- | --- |
| `maxRows` | 500 | Rows a query may return. The last `EVALUATE` is wrapped in `TOPN(maxRows + 1, …)` so the model computes no more; extra rows are cut and the AI is told ("Cut at the app's limit"). |
| `timeoutMs` | 5000 | A query is aborted after this time; the AI is asked to make it lighter. |
| `concurrency` | 1 | Queries running at the same time on one model. |
| `intervalMs` | 300 | Minimum delay between two queries on one model. |
| `cacheMs` | 120000 | The same query on the same model within this delay reuses the rows already read. |
| `maxEmbeds` | 4 | Power BI pages kept loaded by the Presenter (each page is embedded once and reused). |

## Practices

- **Return few rows.** The `executeQueries` API returns at most 100,000 rows or 1,000,000 values per
  query and is limited to 120 queries per minute per user: the app's row limit stays well below, and
  the AI knows it (tool descriptions) and can count rows first with
  `EVALUATE ROW("n", COUNTROWS(<table>))`.
  [Datasets - Execute Queries](https://learn.microsoft.com/rest/api/power-bi/datasets/execute-queries)
- **Aggregate in the model, not in the app.** `SUMMARIZECOLUMNS` rather than `FILTER` on a whole
  table or reading raw rows; the `query` tool writes these queries from a spec.
  [DAX: Avoid using FILTER as a filter argument](https://learn.microsoft.com/power-bi/guidance/dax-avoid-avoid-filter-as-filter-argument),
  [SUMMARIZECOLUMNS](https://learn.microsoft.com/dax/summarizecolumns-function-dax)
- **Stop long queries.** A query that runs for seconds blocks the capacity for others: aborted after
  `timeoutMs`. [Optimization guide for Power BI](https://learn.microsoft.com/power-bi/guidance/power-bi-optimization)
- **Few queries at a time, and not twice.** One query at a time per model, spaced, and a short cache:
  the shared (Pro) capacity is not flooded when the AI checks several figures.
  [Power BI REST API limitations](https://learn.microsoft.com/power-bi/developer/embedded/rest-api-limitations)
- **Refreshes stay with the user.** About 8 refreshes a day per model on Pro: started from the UI only,
  never by the AI. [Data refresh in Power BI](https://learn.microsoft.com/power-bi/connect-data/refresh-data)
- **Embed a report page once.** The Presenter loads each Power BI page once and reuses it for every
  slide showing it (a visual is a zone of that page), instead of one embed per slide.
  [Power BI embedded analytics best practices](https://learn.microsoft.com/power-bi/developer/embedded/embedded-performance-best-practices)
