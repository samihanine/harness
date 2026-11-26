/**
 * Teams through Microsoft Graph, as the signed-in user (Chat.ReadWrite of the Office sign-in): the user's chats
 * (group and one-to-one), a chat's last message (one light call, every few seconds), its messages, their
 * pictures, and sending a message (with a GIF as hosted content). Replaceable for the checks (a fake chat).
 */
import { graph, request } from "../ms/api";

export type TeamsMessage = { id: string; createdDateTime: string; messageType: string; deletedDateTime?: string | null; from?: { user?: { id: string; displayName: string } | null } | null; body: { contentType: string; content: string }; attachments?: { contentType: string; name?: string; contentUrl?: string }[]; mentions?: { mentioned?: { user?: { displayName?: string } } }[] };
export type TeamsChat = { id: string; topic?: string | null; chatType: string; members?: { displayName?: string }[]; lastMessagePreview?: { id: string; createdDateTime: string; body?: { content?: string } } | null };

export type TeamsApi = {
  me(): Promise<{ id: string; displayName: string }>;
  chats(): Promise<TeamsChat[]>;
  lastMessage(chatId: string): Promise<{ id: string; createdDateTime: string } | undefined>;
  messages(chatId: string, top: number): Promise<TeamsMessage[]>;
  pictures(chatId: string, message: TeamsMessage): Promise<{ mediaType: string; data: string }[]>;
  send(chatId: string, html: string, gif?: Uint8Array): Promise<string>;
};

const graphApi: TeamsApi = {
  me: () => graph("/me?$select=id,displayName"),
  async chats() {
    const r = await graph(`/me/chats?$expand=members,lastMessagePreview&$top=50`);
    return (r.value as TeamsChat[]).filter((c) => c.chatType === "group" || c.chatType === "oneOnOne").sort((a, b) => (b.lastMessagePreview?.createdDateTime ?? "").localeCompare(a.lastMessagePreview?.createdDateTime ?? ""));
  },
  async lastMessage(chatId) {
    const r = await graph(`/chats/${chatId}?$expand=lastMessagePreview`);
    return r.lastMessagePreview ?? undefined;
  },
  async messages(chatId, top) {
    const r = await graph(`/chats/${chatId}/messages?$top=${Math.min(50, top)}&$orderby=createdDateTime desc`);
    return r.value as TeamsMessage[];
  },
  async pictures(chatId, m) {
    // Pictures pasted in a message are its hosted contents (their links are in its HTML).
    const ids = [...m.body.content.matchAll(/hostedContents\/([^/"]+)\/\$value/g)].map((x) => x[1]);
    const out: { mediaType: string; data: string }[] = [];
    for (const id of ids.slice(0, 4)) {
      const r = await request("graph", `/chats/${chatId}/messages/${m.id}/hostedContents/${id}/$value`);
      if (r.ok) out.push({ mediaType: r.headers.get("content-type") ?? "image/png", data: Buffer.from(await r.arrayBuffer()).toString("base64") });
    }
    return out;
  },
  async send(chatId, html, gif) {
    const body = gif
      ? { body: { contentType: "html", content: `${html}<p><img src="../hostedContents/1/$value" width="160" height="160" alt="gif"></p>` }, hostedContents: [{ "@microsoft.graph.temporaryId": "1", contentBytes: Buffer.from(gif).toString("base64"), contentType: "image/gif" }] }
      : { body: { contentType: "html", content: html } };
    const r = await graph(`/chats/${chatId}/messages`, "POST", body);
    return r.id as string;
  },
};

let api: TeamsApi = graphApi;
export const teamsApi = () => api;
/** A fake Teams for the checks. */
export const useTeamsApi = (fake: TeamsApi) => void (api = fake);

/** A message's text (its HTML without tags). */
export const messageText = (m: TeamsMessage) =>
  m.body.content
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li)>/gi, "\n")
    .replace(/<img[^>]*>/gi, "[picture]")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

/** A chat's name: its topic, else its members. */
export const chatName = (c: TeamsChat, me?: string) => c.topic || (c.members ?? []).map((m) => m.displayName).filter((n) => n && n !== me).join(", ") || "Chat";
