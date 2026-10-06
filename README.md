# harness

Three independent apps (each folder has its own dependencies and runs alone):

| App | Port | |
| --- | --- | --- |
| [ia-agent](ia-agent) | 3200 | The AI agent panel (100% front-end), embedded by the others in an iframe and connected to them with MCP. |
| [excel-agent](excel-agent) | 3202 | Local Excel files edited through schemas (table / form). |
| [pbi-agent](pbi-agent) | 3201 | Power BI: database of reports / datasets / guides, DAX runner, research, builder, viewer. Has a small local API (sign-in). |

```bash
cd ia-agent && bun install && bun dev        # always needed for the agent panel
cd excel-agent && bun install && bun dev
cd pbi-agent && bun install && bun dev
```

Same stack everywhere: Bun, Vite, React 19, TanStack Router, Tailwind v4, shadcn (only the components used).
