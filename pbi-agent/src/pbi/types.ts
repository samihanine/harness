/** Normalized content of reports and semantic models (stored as content_json). */

/** "Table[Column]", "[Measure]" or "Sum(Table[Column])". */
export type FieldRef = string;

export type VisualContent = {
  name: string;
  type: string;
  title?: string;
  x: number;
  y: number;
  z?: number;
  width: number;
  height: number;
  hidden?: boolean;
  /** Fields by data role ("Category", "Y", "Values", "Rows"…). */
  fields: Record<string, FieldRef[]>;
  filters: unknown[];
  /** Formatting objects as Power BI stores them (objects / vcObjects). */
  objects?: unknown;
  containerObjects?: unknown;
  /** Raw visual definition (as found in the file / API). */
  raw?: unknown;
};

export type PageContent = {
  name: string;
  displayName: string;
  width: number;
  height: number;
  hidden?: boolean;
  filters: unknown[];
  visuals: VisualContent[];
};

export type ReportContent = {
  /** How it was read: "pbix" (full definition) or "embed" (what the JavaScript API exposes). */
  source: "pbix" | "pbir" | "embed";
  pages: PageContent[];
  filters: unknown[];
  theme?: unknown;
  warnings: string[];
  extractedAt: string;
};

export type ColumnInfo = {
  name: string;
  dataType?: string;
  description?: string;
  displayFolder?: string;
  hidden?: boolean;
  /** From COLUMNSTATISTICS when available. */
  stats?: { min?: unknown; max?: unknown; cardinality?: number };
};

export type MeasureInfo = {
  name: string;
  expression?: string;
  formatString?: string;
  description?: string;
  displayFolder?: string;
  hidden?: boolean;
};

export type TableInfo = { name: string; description?: string; hidden?: boolean; columns: ColumnInfo[]; measures: MeasureInfo[] };

export type ModelContent = {
  tables: TableInfo[];
  relationships: { from: string; to: string; active: boolean; crossFilter?: string }[];
  /** What could not be read (missing permissions…). */
  warnings: string[];
  /** Methods that worked, e.g. ["INFO.VIEW.TABLES", "COLUMNSTATISTICS"]. */
  methods: string[];
  extractedAt: string;
};

/** Connection of a report or dataset (config_json). */
export type Connection =
  | { kind: "remote"; datasetId: string; workspaceId?: string; datasetName?: string }
  | { kind: "excel"; fileName: string; sheets: string[] };
