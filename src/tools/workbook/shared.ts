/** The workbook of the screen (an .xlsx in the user's OneDrive), and Excel's addresses. */
import type { DatasetRef } from "@shared/types";
import type { ToolContext } from "../tool";

export type WorkbookDoc = { itemId: string; title?: string; webUrl?: string };

export function workbookOf(ctx: ToolContext) {
  const d = ctx.scope.doc as WorkbookDoc | undefined;
  if (!d?.itemId) throw new Error("No workbook on the screen.");
  return d.itemId;
}

/** A model of the screen by name or id. */
export function modelOf(ctx: ToolContext, given?: string): DatasetRef {
  const list = ctx.scope.datasets ?? [];
  const m = given ? list.find((d) => d.id === given || d.name.toLowerCase() === given.toLowerCase()) : list[0];
  if (!m) throw new Error(given ? `No model "${given}". Models: ${list.map((d) => `"${d.name}"`).join(", ") || "none (check them in Semantic models)"}.` : "No model checked for this analysis (Semantic models tab).");
  return m;
}

export const letters = (n: number): string => (n > 26 ? letters(Math.floor((n - 1) / 26)) : "") + String.fromCharCode(65 + ((n - 1) % 26));
export const colIndex = (l: string) => l.toUpperCase().split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);

/** "B3" → the address of a block of `rows` × `cols` from it ("B3:E10"). */
export function blockAt(at: string, rows: number, cols: number) {
  const m = at.trim().replace(/\$/g, "").match(/^([A-Z]+)(\d+)$/i);
  if (!m) throw new Error(`"${at}" is not a cell address (like B3).`);
  const c = colIndex(m[1]);
  const r = Number(m[2]);
  return `${m[1].toUpperCase()}${r}:${letters(c + Math.max(1, cols) - 1)}${r + Math.max(1, rows) - 1}`;
}

/** Rows as a compact text table. */
export function gridText(values: unknown[][], formulas?: unknown[][], startRow = 1, startCol = 1) {
  return values
    .map((row, r) => `${startRow + r}: ${row.map((v, c) => {
      const f = formulas?.[r]?.[c];
      const shown = v === "" || v === null ? "·" : JSON.stringify(v);
      return typeof f === "string" && f.startsWith("=") ? `${letters(startCol + c)}=${shown} {${f}}` : `${letters(startCol + c)}=${shown}`;
    }).join(" | ")}`)
    .join("\n");
}
