/**
 * Storage layer: where a file's rows and images live. The app (lib/excel.ts) only talks to
 * these interfaces; local files (File System Access + ExcelJS) and SharePoint / OneDrive
 * (Microsoft Graph Excel API) implement them.
 */

/** Plain cell value exchanged with a backend (dates as "YYYY-MM-DD" or Excel serial numbers). */
export type Cell = string | number | boolean | Date | null;
export type Values = Record<string, Cell>;

export type ColumnSpec = {
  name: string;
  /** Dropdown list of allowed values. */
  list?: string[];
  date?: boolean;
  /** Rich text: values are canonical HTML (see lib/rich.ts), stored as rich runs when the backend can. */
  rich?: boolean;
  wrap?: boolean;
  width?: number;
  note?: string;
  /** Cell colours of option values (hex "#rrggbb"), applied by conditional formatting. */
  colors?: { value: string; fill: string; font: string }[];
  /** Several values per cell ("a; b"): colours match contained values, no dropdown. */
  multiple?: boolean;
  /** Computed column: IMAGE() of the link in that column (never written by rows). */
  imageOf?: string;
};

/**
 * Formula of a picture cell, from the cell holding the link (e.g. "D5"): only links can be shown
 * (local images are file names), and SharePoint sharing links open a page unless "download=1"
 * asks for the file itself. Files store IMAGE() as _xlfn.IMAGE(); Graph takes the plain name.
 */
export const imageFormula = (link: string, stored: boolean) => {
  const url = `IF(ISNUMBER(SEARCH("sharepoint.com",${link})),${link}&IF(ISNUMBER(SEARCH("?",${link})),"&","?")&"download=1",${link})`;
  return `IF(LEFT(${link},4)="http",${stored ? "_xlfn." : ""}IMAGE(${url}),"")`;
};

/** Height of data rows (points): taller when the sheet shows images. */
export const rowHeight = (specs: ColumnSpec[]) => (specs.some((s) => s.imageOf) ? 60 : 24);

/** A table of rows (header on top, "id" column first) in one sheet of a workbook. */
export interface TableBackend {
  /**
   * Opens (or reopens) the sheet: created with the columns when missing, missing columns added,
   * rows without id given one. Returns the rows by column name.
   */
  open(sheet: string, columns: ColumnSpec[]): Promise<Values[]>;
  /** Changes are queued (applied in order) and persisted by flush(). */
  update(changes: { id: string; values: Values }[]): void;
  append(rows: Values[]): void;
  remove(ids: string[]): void;
  flush(): Promise<void>;
  /** True when the data was changed by someone else since the last open / flush. */
  changedOutside(): Promise<boolean>;
  /** How often changedOutside() is worth checking (ms). */
  readonly pollEvery: number;
}

/** Where images are stored; cells hold a file name (local) or a link (SharePoint, web). */
export interface ImageStore {
  readonly label: string;
  /** Saves a file and returns the value to put in the cell. */
  save(file: File): Promise<string>;
  /** Existing images: `pick(id)` returns the cell value for one of them. */
  list(): Promise<{ id: string; name: string }[]>;
  pick(id: string): Promise<string>;
  /** Displayable URL of a cell value of this store. */
  resolve(value: string): Promise<string | null>;
}

/** A file or folder on SharePoint / OneDrive. */
export type DriveRef = { driveId: string; itemId: string; name: string; webUrl: string };

export type FileSource = ({ kind: "local" } & { name: string }) | ({ kind: "sharepoint" } & DriveRef);
export type ImagesSource = ({ kind: "local" } & { name: string }) | ({ kind: "sharepoint" } & DriveRef);
