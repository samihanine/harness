/** Types shared by the browser and the server: what a chat request carries and the events it streams back. */

/** A Power BI item: `appId` when opened from a Power BI app, `groupId` its workspace (none: My workspace). */
export type Ref = { id: string; groupId?: string; appId?: string };

export type ScreenId = "trainer" | "builder" | "scientist" | "presenter";

/**
 * A report as the app keeps it. `access`: "edit" when its PBIR definition can be read and written (the
 * Builder writes it), "editor" when it can be edited only through the Power BI editor (older PBIR-Legacy
 * format, or a definition the API refuses: the Builder works in the embedded editor), "view" otherwise.
 * `structure`: its structure read once (pages, visuals, fields), refreshed on demand.
 */
export type ReportAccess = "view" | "edit" | "editor";
export type ReportRef = Ref & { name: string; datasetId?: string; datasetGroupId?: string; access?: ReportAccess; webUrl?: string; structure?: unknown; structureAt?: string };
export type DatasetRef = Ref & { name: string; access?: "view" | "edit" };

/** A project: a name and a Confluence folder whose pages are the context of the conversations. */
export type Project = { id: string; name: string; folderId: string };

/** A file sent to the server (image for the AI, Excel / CSV for the Data Scientist, PowerPoint for the Presenter). */
export type FileRef = { id: string; name: string; type: string; size: number };

/**
 * What the screen works on, sent with every message: the server reads nothing from the browser's storage.
 * `doc`: the document the AI edits on this screen (an analysis, a deck, a page draft), the browser's copy being the reference.
 */
export type ChatScope = {
  report?: ReportRef;
  dataset?: DatasetRef;
  /** Reports the Builder takes as inspiration. */
  references?: ReportRef[];
  /** Models and files the Data Scientist works with. */
  datasets?: DatasetRef[];
  files?: FileRef[];
  /** Projects chosen in the chat: their Confluence pages are the conversation's context. */
  projects?: string[];
  /** Structures and models the browser already holds (read once, refreshed on demand): not read again. */
  models?: Record<string, unknown>;
  /** Presenter: "create" (a deck written in markdown) or "edit" (an uploaded PowerPoint). */
  mode?: "create" | "edit";
  doc?: unknown;
};

export type AskQuestion = { question: string; options: string[]; multiple?: boolean };

/** Streamed by POST /api/chat (one JSON per SSE message). */
export type ChatEvent =
  | { type: "start"; runId: string; model: string }
  | { type: "text"; text: string }
  | { type: "status"; label: string }
  | { type: "tool"; id: string; name: string; args: unknown; status: "running" | "done" | "error"; result?: string; images?: number; ms?: number }
  /** A tool the browser runs (it holds the embedded report): answered with POST /api/chat/tool-result. */
  | { type: "client_tool"; callId: string; name: string; args: unknown }
  | { type: "ask"; message?: string; questions: AskQuestion[] }
  /** Something changed that the screen shows: `doc` (the new document), `reload` (the report was saved)… */
  | { type: "effect"; name: string; data?: unknown }
  | { type: "done"; usage?: { input: number; output: number } }
  | { type: "error"; error: string };

/** What the browser keeps of a conversation (IndexedDB) and shows. */
export type ChatItem =
  | { kind: "user"; text: string; files?: FileRef[] }
  | { kind: "assistant"; text: string; tools: { id: string; name: string; args: unknown; status: "running" | "done" | "error"; result?: string; images?: number; ms?: number }[]; ask?: { message?: string; questions: AskQuestion[] }; error?: string; running?: boolean };

export type ChatRequest = {
  conversationId: string;
  screen: ScreenId;
  message: string;
  files?: FileRef[];
  scope: ChatScope;
  /** The conversation so far as text, used when the server no longer holds it (restarted). */
  transcript?: string;
};

/* --------------------------------- Settings -------------------------------- */

export type Provider = "anthropic" | "maia" | "claude-code";

/** Public view of the settings: secrets are never sent to the browser, only whether they are set. */
export type PublicSettings = {
  general: { primaryColor: string; screens: Record<ScreenId, boolean>; mascot: boolean; language?: string };
  ai: {
    provider: Provider;
    model: string;
    effort: "low" | "medium" | "high";
    maxSteps: number;
    anthropic: { keySet: boolean };
    maia: { url: string; clientIdSet: boolean; clientSecretSet: boolean; maison: string; useCaseId: string; environment: string };
    claudeCode: { baseUrl: string; headers: string; keyHelper: string; keySet: boolean };
  };
  powerbi: { maxRows: number; timeoutMs: number; concurrency: number; cacheMs: number };
  confluence: { url: string; email: string; folderId: string; keySet: boolean; projects: Project[] };
  /** Settings given by the .env file (the values a reset goes back to). */
  env: string[];
};

/** A change of the settings from the browser: secrets may be written, never read. */
export type SettingsPatch = {
  general?: Partial<PublicSettings["general"]>;
  ai?: Partial<Omit<PublicSettings["ai"], "anthropic" | "maia" | "claudeCode">> & {
    anthropicKey?: string;
    maia?: Partial<{ url: string; clientId: string; clientSecret: string; maison: string; useCaseId: string; environment: string }>;
    claudeCode?: Partial<{ baseUrl: string; headers: string; keyHelper: string; key: string }>;
  };
  powerbi?: Partial<PublicSettings["powerbi"]>;
  confluence?: Partial<{ url: string; email: string; folderId: string; key: string; projects: Project[] }>;
};

export const SCREENS: ScreenId[] = ["trainer", "builder", "scientist", "presenter"];

/** The project of the Confluence folder of the settings (CONFLUENCE_FOLDER_ID). */
export const DEFAULT_PROJECT = "default";
