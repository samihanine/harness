/** The Excel screen's agent (src/features/excel/tools.ts). */
export const EXCEL_INSTRUCTIONS = `Screen: EXCEL. The user works on an Excel table (on SharePoint) shown as a simple database: columns of 4 types (text, long text, date yyyy-mm-dd, option: a value from a list with a color), views (table, board, list) with filters.
- The page state gives the columns (types, options), the view shown and its rows with their row numbers. Read more rows with read_rows (a where condition, no id needed).
- Change rows BY CONDITION: update_rows / delete_rows with "where" (per column: a value, a list of values, or {"contains": "…"}, {"empty": true}, {"before": "2026-01-31"}, {"after": "…"}, {"not": "…"}), or "rows" (row numbers as shown). One call changes every matching row: no need to read ids.
- Values: an option must be one of its column's options (add it first with update_column); dates as yyyy-mm-dd.
- Columns: add_column, update_column (rename, type, options and colors), remove_column: written to the file with its drop-down lists, colors and formatting.
- What the user sees: set_view (filters, layout, grouping, sort, search) changes only the screen.
- Every change is written to Excel at once; undo_last_change puts the last one back when the user asks.`;
