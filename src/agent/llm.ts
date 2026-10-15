/**
 * Deliberately minimal LLM access: plain-text messages in, plain text out.
 * No tool use, no JSON mode: the harness builds everything on top (see harness/).
 *
 * The harness talks to the model through short-lived sessions: one per user request
 * (and one per sub-agent). A session is either a remote conversation (stateful API) or
 * a local message list replayed with every call (stateless API).
 */
import { settings } from "@/lib/store";

/* -------------------------------------------------------------------------- */
/* Provider settings — change these to target another compatible chat API.    */
/* -------------------------------------------------------------------------- */

/** Base URL of the API (called from the browser: it must allow CORS). */
export const AI_BASE_URL = "https://api.openai.com/v1";

/** Exact model names (the first one is the default). */
export const AI_MODELS: readonly string[] = ["gpt-6-luna", "gpt-5.5", "gpt-5", "gpt-5-mini"];

/**
 * Remote conversations: path that creates one (POST, body `{ name }`) and the response field
 * holding its id. `null` = stateless API, the session history is sent with every message.
 */
const AI_CONVERSATION_PATH: string | null = null;
const AI_CONVERSATION_ID_FIELD = "id";

/**
 * Files: path that uploads one (multipart POST: `file` + the fields below) and the response field
 * holding its id. `null` = no file upload.
 */
const AI_FILE_UPLOAD_PATH: string | null = "/files";
const AI_FILE_UPLOAD_FIELDS: Record<string, string> = { purpose: "user_data" };
const AI_FILE_ID_FIELD = "id";

/**
 * Path that adds an uploaded file to a remote conversation (POST; `{conversationId}`, `{fileId}`
 * replaced). `null` = stateless API: the file goes with the next message of the conversation.
 */
const AI_CONVERSATION_FILE_PATH: string | null = null;

/** Path that deletes a remote conversation, called when the user deletes the conversation in the app (`null` = none). */
const AI_DELETE_PATH: string | null = null;

/** Path of a completion; `{conversationId}` is replaced by the remote conversation id. */
const AI_MESSAGE_PATH = "/chat/completions";

/** "json": one JSON response. "sse": event stream (`event: content` chunks until `event: done`). */
const AI_RESPONSE: "json" | "sse" = "json";

/**
 * Request headers built from the saved AI key. The key is either a plain token or several
 * `NAME=value` pairs (one per line), available in `fields`.
 */
const AI_HEADERS = (key: AiKey): Record<string, string> => ({ Authorization: `Bearer ${key.token}` });

/**
 * Request body. `messages` = whole session (stateless APIs), `message` = the new user turn, `files` =
 * ids of the files added to the conversation since the last turn (they go with this message).
 */
const AI_BODY = ({ model, messages }: { model: string; messages: ChatMessage[]; message: string; files: string[] }) => ({
  model,
  messages: messages.map((m) => (m.files?.length ? { role: m.role, content: [{ type: "text", text: m.content }, ...m.files.map((id) => ({ type: "file", file: { file_id: id } }))] } : { role: m.role, content: m.content })),
});

/** Reply text of a "json" response. */
const AI_REPLY = (data: { choices?: { message?: { content?: string } }[] }) =>
  data.choices?.[0]?.message?.content ?? "";

/* -------------------------------------------------------------------------- */

export type ChatMessage = { role: "user" | "assistant"; content: string; files?: string[] };
type AiKey = { token: string; fields: Record<string, string> };

/** "abc" → token; "NAME=value" lines → fields (the first value is also the token). */
function parseKey(raw: string): AiKey {
  const fields = Object.fromEntries([...raw.matchAll(/([A-Za-z0-9_]+)=(\S+)/g)].map((m) => [m[1], m[2]]));
  return { token: Object.values(fields)[0] ?? raw.trim(), fields };
}

export class AiAuthError extends Error {
  constructor(status: number) {
    super(status ? `The AI key was rejected (HTTP ${status}).` : "Missing AI key: add it in the AI settings (⚙ at the top right).");
  }
}

/** A conversation with the model. `chars` measures what it holds (to restart it before it overflows). */
export type Session = {
  send(text: string, signal?: AbortSignal): Promise<string>;
  /** Id of the remote conversation (stateful APIs), kept with the app conversation to delete it later. */
  remote(): Promise<string | undefined>;
  readonly chars: number;
};

/** Files added to each conversation and not sent yet (they go with its next message). */
const pendingFiles = new Map<string, string[]>();

/** Uploads a file to the AI API; returns its id (to add it to a conversation). */
export async function uploadAiFile(file: File) {
  if (!AI_FILE_UPLOAD_PATH) throw new Error("This AI API does not take files.");
  const form = new FormData();
  form.append("file", file);
  for (const [k, v] of Object.entries(AI_FILE_UPLOAD_FIELDS)) form.append(k, v);
  const data = await (await request("POST", AI_FILE_UPLOAD_PATH, form)).json();
  const id = data[AI_FILE_ID_FIELD];
  if (!id) throw new Error("The AI API did not return a file id");
  return String(id);
}

/**
 * Adds an uploaded file to a conversation (the app conversation's id, as given to openSession): sent
 * with its next message, or added to its remote conversation when the API has a path for it.
 */
export async function addFileToConversation(conversationId: string, fileId: string) {
  const pending = pendingFiles.get(conversationId) ?? [];
  if (!pending.includes(fileId)) pendingFiles.set(conversationId, [...pending, fileId]);
}

/** `conversationId`: the app conversation, for the files added to it (see addFileToConversation). */
export function openSession(model: string, name: string, conversationId?: string): Session {
  const messages: ChatMessage[] = [];
  let remoteId: Promise<string> | undefined;
  let chars = 0;
  return {
    get chars() {
      return chars;
    },
    async send(text, signal) {
      if (AI_CONVERSATION_PATH) remoteId ??= createRemote(name, signal);
      const id = remoteId ? await remoteId : "";
      const files = (conversationId && pendingFiles.get(conversationId)) || [];
      if (conversationId) pendingFiles.delete(conversationId);
      // A stateful API takes the files into its conversation; a stateless one gets them with this message.
      if (AI_CONVERSATION_FILE_PATH && id) for (const f of files) await request("POST", AI_CONVERSATION_FILE_PATH.replace("{conversationId}", id).replace("{fileId}", f), {}, signal);
      const sent: ChatMessage = { role: "user", content: text, ...(!AI_CONVERSATION_FILE_PATH && files.length && { files }) };
      const body = AI_BODY({ model, messages: [...messages, sent], message: text, files });
      const response = await request("POST", AI_MESSAGE_PATH.replace("{conversationId}", id), body, signal);
      const reply = AI_RESPONSE === "sse" ? await readStream(response) : String(AI_REPLY(await response.json()));
      // Stateless APIs need the history on our side; stateful ones keep it remotely.
      if (!AI_CONVERSATION_PATH) messages.push(sent, { role: "assistant", content: reply });
      chars += text.length + reply.length;
      return reply;
    },
    async remote() {
      return remoteId ? remoteId.catch(() => undefined) : undefined;
    },
  };
}

/** Deletes a remote conversation (when the API has a delete path). */
export async function deleteRemoteConversation(id: string) {
  if (AI_DELETE_PATH && id) await request("DELETE", AI_DELETE_PATH.replace("{conversationId}", id));
}

/** Calls the API from the browser, with the key saved in the AI settings. */
async function request(method: string, path: string, body?: unknown, signal?: AbortSignal) {
  const { aiKey } = await settings.get();
  if (!aiKey) throw new AiAuthError(0);
  // A file upload (FormData) sets its own content type.
  const form = body instanceof FormData;
  const response = await fetch(`${AI_BASE_URL}${path}`, {
    method,
    signal,
    headers: { ...(form ? {} : { "Content-Type": "application/json" }), ...AI_HEADERS(parseKey(aiKey)) },
    body: body === undefined ? undefined : form ? body : JSON.stringify(body),
  });
  if (response.status === 401 || response.status === 403) throw new AiAuthError(response.status);
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error?.message ?? data.content ?? `AI request failed (${response.status})`);
  }
  return response;
}

async function createRemote(name: string, signal?: AbortSignal) {
  const data = await (await request("POST", AI_CONVERSATION_PATH!, { name }, signal)).json();
  const id = data[AI_CONVERSATION_ID_FIELD];
  if (!id) throw new Error("The AI API did not return a conversation id");
  return String(id);
}

/** Server-sent events: concatenates `content` events until `done` (a done payload `{"valid":false}` is an error). */
async function readStream(response: Response) {
  if (!response.body) throw new Error("Empty stream from the AI API");
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let text = "";
  let event = "";
  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split("\n");
    buffer = lines.pop() ?? "";
    for (const line of lines) {
      if (line.startsWith("event:")) event = line.slice(6).trim();
      if (!line.startsWith("data:")) continue;
      const data = line.slice(5).replace(/^ /, "");
      if (event === "content") text += data;
      if (event === "done") {
        if (/"valid"\s*:\s*false/.test(data)) throw new Error("The AI API returned an invalid completion");
        return text.trim();
      }
    }
  }
  return text.trim();
}
