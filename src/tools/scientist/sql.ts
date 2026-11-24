/** SQL on the analysis's tables: to explore, and steps (views) that combine sources. */
import { z } from "zod";
import { ident, rowsText, safeName, select } from "@server/data/engine";
import { tool } from "../tool";
import { analysisOf, engine, saveAnalysis } from "./shared";

export const sqlQuery = tool({
  name: "sql",
  description: "Run a SELECT (SQLite: joins, GROUP BY, CTEs, window functions, CASE, strftime for dates stored as text 'YYYY-MM-DD') on the analysis's tables and steps. To explore and check figures.",
  args: z.object({ query: z.string(), max_rows: z.number().optional() }),
  async run(a, ctx) {
    const r = select(await engine(ctx), a.query, Math.min(500, a.max_rows ?? 100));
    return `${rowsText(r.columns, r.rows, 100)}${r.truncated ? "\n(more rows exist)" : ""}`;
  },
});

export const createStep = tool({
  name: "create_step",
  description: "Save a SQL step of the analysis: a named view (CREATE VIEW name AS <select>) other queries and blocks can use, e.g. a join of a model's sales with an uploaded target file. Replaces a step of the same name.",
  args: z.object({ name: z.string(), sql: z.string().describe("A SELECT."), description: z.string().optional() }),
  changes: true,
  async run(a, ctx) {
    const e = await engine(ctx);
    const name = safeName(a.name);
    select(e, `SELECT * FROM (${a.sql.replace(/;+\s*$/, "")}) LIMIT 1`); // checked first
    e.db.run(`DROP VIEW IF EXISTS ${ident(name)}`);
    e.db.run(`CREATE VIEW ${ident(name)} AS ${a.sql.replace(/;+\s*$/, "")}`);
    const an = analysisOf(ctx);
    saveAnalysis(ctx, { ...an, steps: [...an.steps.filter((s) => s.name !== name), { name, sql: a.sql, description: a.description }] });
    const sample = select(e, `SELECT * FROM ${ident(name)} LIMIT 5`);
    return `Step ${name} saved.\n${rowsText(sample.columns, sample.rows)}`;
  },
});
