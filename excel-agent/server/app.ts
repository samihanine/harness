/** Mini API next to the front-end (same origin): Microsoft sign-in and tokens for Graph. */
import { Hono } from "hono";
import { accessToken, logout, startLogin, status } from "./auth";

export const app = new Hono().basePath("/api");

// Same-origin only: the API hands out tokens.
app.use("*", async (c, next) => {
  const origin = c.req.header("origin");
  if (origin && new URL(origin).host !== new URL(c.req.url).host) return c.json({ error: "Forbidden origin" }, 403);
  await next();
});

app.get("/status", async (c) => c.json(await status()));
app.post("/login", async (c) => c.json(await startLogin()));
app.post("/logout", async (c) => {
  await logout();
  return c.json({ ok: true });
});
app.get("/token", async (c) => {
  const token = await accessToken("graph");
  return token ? c.json(token) : c.json({ error: "login_required" }, 401);
});
