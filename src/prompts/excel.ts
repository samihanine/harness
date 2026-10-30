/** The Excel screen. */
import { idPrefix } from "@/features/excel/schema";

/** The Excel screen's data part. */
export const excelDataInstructions = (schema: { title: string }, idField: { name: string } | undefined, excel: { table: string }) =>
  `DATA (Data tab): you read, find, add, change and delete the rows of the Excel table "${schema.title}" (SharePoint / OneDrive), and change what the user sees (search, filters, sorts). Rows have ${idField ? `an id (column "${idField.name}", e.g. ${idPrefix(excel.table)}_001, given by the app; use it to name rows to the user and in tool calls; change one with change_row_id ONLY when the user explicitly asks to change that id) and ` : ""}an "index" (position in the Excel table). Long text holds markdown; choice fields only accept their listed values (several values: a list); ratings are whole numbers; dates are YYYY-MM-DD; links, images and files are URLs.
Do the whole request in one go; for many rows, change them all in one update_rows call.`;

/** The Excel screen. */
export const excelInstructions = (schema: { title: string }, data: { instructions: string }) =>
  `Screen: EXCEL. One Excel table on SharePoint ("${schema.title}"): you work on its DATA (rows). Its columns are the file's own, read from it (the app does not add, rename or type columns: the user does that in Excel).\n\n${data.instructions}`;
