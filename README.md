# Hub

Power BI, Excel analyses and PowerPoint decks, with AI assistants that explore, ask, then act.

- **Trainer**: ask about a report; the assistant reads it, sets the right slicers and filters, explains.
- **Builder**: create and change reports (PBIR) from a sentence, a screenshot, or reference reports.
- **Data Scientist**: combine semantic models (DAX) with Excel / CSV files in SQL; analyses exported to Excel.
- **Presenter**: decks from the template written in markdown, or an existing PowerPoint updated in place
  (Power BI screenshots refreshed).
- **Reports, Semantic models, Confluence**: structures (JSON), DAX, documentation used as context.

## Run

```bash
bun install
bun dev            # app http://localhost:3300, API :3301
```

Sign in to Microsoft in Settings › Power BI, choose the AI in Settings › AI (Maia, Claude Code or an
Anthropic key; `CLAUDE_API_KEY` / `ANTHROPIC_API_KEY`, `MAIA_CLIENT_ID` / `MAIA_CLIENT_SECRET` and
`CONFLUENCE_API_URL` / `CONFLUENCE_API_KEY` are read at first start).

Production: `bun run build && bun run start` (one port, 3300), or `bun run export:server` (one executable;
`HOST` and `HUB_ALLOWED_HOSTS` to serve a network). Pictures of reports need Edge or Chrome on the machine.
