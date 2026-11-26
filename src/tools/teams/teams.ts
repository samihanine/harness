/**
 * The Teams assistant's own tools: earlier messages of the chat and their pictures, a data table shown in its
 * answer like a pivot table (the model, its filters, rows, columns and values, then the figures), and a GIF of
 * the mascot sent with the answer.
 */
import { z } from "zod";
import { dax } from "@server/pbi/dax";
import { messageText, teamsApi } from "@server/teams/api";
import { MOODS } from "@server/teams/moods";
import { modelOf } from "../helpers";
import { tool, type ToolContext } from "../tool";

const chatOf = (ctx: ToolContext) => {
  const id = (ctx.scope.doc as { chatId?: string } | undefined)?.chatId;
  if (!id) throw new Error("No Teams chat.");
  return id;
};

export const readChat = tool({
  name: "read_chat",
  description: "Earlier messages of the Teams chat (the latest first is not needed: you already have the last ones): up to 50, oldest first, with their ids.",
  args: z.object({ count: z.number().int().min(1).max(50).optional() }),
  async run(a, ctx) {
    const list = (await teamsApi().messages(chatOf(ctx), a.count ?? 30)).filter((m) => m.messageType === "message").reverse();
    return list.map((m) => `(${m.id}) [${m.createdDateTime.slice(0, 16).replace("T", " ")}] ${m.from?.user?.displayName ?? "?"}: ${messageText(m).slice(0, 1500)}${m.attachments?.length ? ` [attachments: ${m.attachments.map((x) => x.name ?? x.contentType).join(", ")}]` : ""}`).join("\n");
  },
});

export const lookAtMessage = tool({
  name: "look_at_message",
  description: "See the pictures pasted in a message of the chat (by its id): a screenshot of a report, a chart…",
  args: z.object({ message_id: z.string() }),
  async run(a, ctx) {
    const api = teamsApi();
    const m = (await api.messages(chatOf(ctx), 50)).find((x) => x.id === a.message_id);
    if (!m) throw new Error(`No message ${a.message_id} among the last 50.`);
    const images = await api.pictures(chatOf(ctx), m);
    return images.length ? { text: `${images.length} picture(s) of the message.`, images } : "This message holds no picture.";
  },
});

const col = (s: string) => {
  const m = s.trim().match(/^'?(.+?)'?\[(.+)\]$/);
  if (!m) throw new Error(`"${s}" is not a column: write 'Table'[Column].`);
  return `'${m[1].replace(/'/g, "''")}'[${m[2].replace(/]/g, "]]")}]`;
};
const measure = (s: string) => `[${s.trim().replace(/^.*?\[(.+)\]$/, "$1").replace(/]/g, "]]")}]`;
const esc = (s: unknown) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;");
const num = (v: unknown) => (typeof v === "number" ? new Intl.NumberFormat("fr-FR", { maximumFractionDigits: Math.abs(v) < 10 ? 2 : 0 }).format(v) : esc(v));

export const dataTable = tool({
  name: "data_table",
  description:
    "A table of figures from a model for the answer, shown like a pivot table: a line with the model, the filters, the rows, the columns and the values, then the figures. Use it whenever the question is about data that can be checked (not for general questions the documentation answers). rows: columns ('Table'[Column]); columns: one column whose values become the table's columns (first value only); values: measures; filters: columns and the values kept. Returns {{table:N}} to put in your answer where the table goes, and the figures for your text.",
  args: z.object({
    model: z.string().optional(),
    rows: z.array(z.string()).max(3).default([]),
    columns: z.string().optional(),
    values: z.array(z.string()).min(1).max(4),
    filters: z.array(z.object({ column: z.string(), values: z.array(z.union([z.string(), z.number()])) })).optional(),
    top: z.number().int().max(50).optional(),
  }),
  async run(a, ctx) {
    const m = modelOf(ctx, a.model);
    const filters = (a.filters ?? []).map((f) => `TREATAS({${f.values.map((v) => (typeof v === "number" ? v : `"${String(v).replace(/"/g, '""')}"`)).join(", ")}}, ${col(f.column)})`);
    const groups = [...a.rows.map(col), ...(a.columns ? [col(a.columns)] : [])];
    const values = a.values.map((v, i) => `"v${i}", ${measure(v)}`);
    const query = `EVALUATE TOPN(${a.top ?? 25 * (a.columns ? 20 : 1)}, SUMMARIZECOLUMNS(${[...groups, ...filters, ...values].join(", ")})${groups.length ? `, [v0], DESC` : ""})`;
    const rows = await dax({ id: m.id, groupId: m.groupId }, query);
    const label = (s: string) => s.replace(/^'?(.+?)'?\[(.+)\]$/, "$2").replace(/^\[(.+)\]$/, "$1");
    const key = (r: Record<string, unknown>, c: string) => r[label(c)] ?? r[c];
    let head: string[];
    let body: string[][];
    if (a.columns) {
      // Pivoted: one column per value of `columns`, the first measure in the cells.
      const cols = [...new Set(rows.map((r) => String(key(r, a.columns!))))].slice(0, 12);
      const lines = new Map<string, Record<string, unknown>>();
      for (const r of rows) {
        const k = a.rows.map((x) => String(key(r, x))).join(" · ") || "Total";
        lines.set(k, { ...(lines.get(k) ?? {}), [String(key(r, a.columns))]: r.v0 });
      }
      head = [a.rows.map(label).join(" · ") || "", ...cols];
      body = [...lines.entries()].slice(0, a.top ?? 25).map(([k, v]) => [esc(k), ...cols.map((c) => num(v[c]))]);
    } else {
      head = [...a.rows.map(label), ...a.values.map(label)];
      body = rows.slice(0, a.top ?? 25).map((r) => [...a.rows.map((x) => esc(key(r, x))), ...a.values.map((_, i) => num(r[`v${i}`]))]);
    }
    const info = [`📊 <b>${esc(m.name)}</b>`, a.filters?.length ? `Filters: ${a.filters.map((f) => `${esc(label(f.column))} = ${esc(f.values.join(", "))}`).join("; ")}` : "", a.rows.length ? `Rows: ${a.rows.map((r) => esc(label(r))).join(", ")}` : "", a.columns ? `Columns: ${esc(label(a.columns))}` : "", `Values: ${a.values.map((v) => esc(label(v))).join(", ")}`].filter(Boolean).join(" · ");
    const th = (c: string, i: number) => `<th style="text-align:${i ? "right" : "left"};padding:2px 8px;border-bottom:1px solid #999">${esc(c)}</th>`;
    const td = (c: string, i: number) => `<td style="text-align:${i ? "right" : "left"};padding:2px 8px">${c}</td>`;
    const html = `<p style="font-size:12px;color:#666">${info}</p><table><tr>${head.map(th).join("")}</tr>${body.map((r) => `<tr>${r.map(td).join("")}</tr>`).join("")}</table>`;
    const n = ((ctx.conversation.scratch.tables as number) ?? 0) + 1;
    ctx.conversation.scratch.tables = n;
    ctx.effect("table", { n, html });
    return `Put {{table:${n}}} in your answer where the table goes. Its figures:\n${[head.join(" | "), ...body.map((r) => r.join(" | "))].join("\n")}`;
  },
});

export const sendGif = tool({
  name: "send_gif",
  description: `Send a GIF of the mascot with your answer, in a mood: ${MOODS.join(", ")}. For a remark outside your job (a compliment, a teasing, an insult: answer with humour, in character), or at the end of an exchange when the person says they have their answer.`,
  args: z.object({ mood: z.enum(MOODS as [string, ...string[]]) }),
  async run(a, ctx) {
    ctx.effect("gif", a.mood);
    return `The ${a.mood} GIF goes with your answer.`;
  },
});

export const TEAMS_TOOLS = [readChat, lookAtMessage, dataTable, sendGif];
