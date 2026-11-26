/** What the Microsoft sign-in allows in Teams: the token's scopes, the user's chats and their last message. */
import { accessToken } from "@server/ms/auth";
import { request } from "@server/ms/api";
const token = await accessToken("graph");
const claims = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
console.log("app:", claims.app_displayname, "scopes:", claims.scp);
const r = await request("graph", "/me/chats?$expand=lastMessagePreview&$top=10");
const body = await r.json();
console.log("chats:", r.status, JSON.stringify(body).slice(0, 600));
