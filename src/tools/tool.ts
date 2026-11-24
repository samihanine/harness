/**
 * Every tool the AI uses is a file of src/tools: its name, what it does (the model reads it), its
 * arguments (zod) and its code. Tools run on the server (`run`); a few act on what the browser shows
 * (the embedded report): those call a browser action with `ctx.browser(...)` (src/tools/browser/).
 */
import { z } from "zod";
import type { ChatEvent, ChatScope, ScreenId } from "@shared/types";

export type ToolImage = { mediaType: string; data: string };
/** A text, or a text with pictures (a screenshot) the model sees. */
export type ToolOutput = string | { text: string; images?: ToolImage[] };

/** What a conversation keeps on the server between messages (drafts, loaded data…), per screen. */
export type ConversationState = {
  id: string;
  screen: ScreenId;
  /** The user was asked a clarification question (required before the first change). */
  asked: boolean;
  /** Free room for the screen's tools (a report draft, the analysis database, the working PowerPoint…). */
  scratch: Record<string, any>;
};

export type ToolContext = {
  scope: ChatScope;
  conversation: ConversationState;
  /** Runs an action in the user's browser (it holds the embedded report) and returns its result. */
  browser: (action: string, args?: Record<string, unknown>) => Promise<any>;
  /** Tells the screen something changed (`doc`: the new document; `reload`: the report was saved…). */
  effect: (name: string, data?: unknown) => void;
  status: (label: string) => void;
  signal: AbortSignal;
};

export type Tool = {
  name: string;
  description: string;
  args: z.ZodObject;
  /** It changes something (a report, a model, a document): refused until the user answered a first question. */
  changes?: boolean;
  run: (args: any, ctx: ToolContext) => Promise<ToolOutput>;
};

export const tool = <S extends z.ZodObject>(t: { name: string; description: string; args: S; changes?: boolean; run: (args: z.infer<S>, ctx: ToolContext) => Promise<ToolOutput> }): Tool => t as Tool;

/** Base64 of bytes. */
export const base64 = (bytes: ArrayBuffer | Uint8Array) => Buffer.from(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)).toString("base64");

/** A JSON value as short text for the model. */
export const json = (value: unknown, max = 60_000) => {
  const text = JSON.stringify(value, null, 1);
  return text.length > max ? `${text.slice(0, max)}\n…[cut at ${max} characters: ask for a smaller part]` : text;
};

export type ChatEmit = (event: ChatEvent) => void;
