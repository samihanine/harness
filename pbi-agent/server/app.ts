/**
 * Mini API next to the front-end (same origin): sign-in, tokens, and a pass-through to the
 * Power BI / Fabric REST APIs for the few endpoints browsers cannot call directly (CORS).
 */
import { Hono } from "hono";
import type { Audience } from "./auth";
import { accessToken, logout, startLogin, status } from "./auth";

const UPSTREAM: Record<string, { base: string; audience: Audience }> = {
  pbi: { base: "https://api.powerbi.com", audience: "powerbi" },
  fabric: { base: "https://api.fabric.microsoft.com", audience: "fabric" },
};

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
  const token = await accessToken((c.req.query("audience") as Audience) ?? "powerbi");
  return token ? c.json(token) : c.json({ error: "login_required" }, 401);
});

/** /api/pbi/v1.0/myorg/… → https://api.powerbi.com/v1.0/myorg/… with the user's token. */
app.all("/:upstream{pbi|fabric}/*", async (c) => {
  const upstream = UPSTREAM[c.req.param("upstream")];
  const token = await accessToken(upstream.audience);
  if (!token) return c.json({ error: "login_required" }, 401);
  const url = new URL(c.req.url);
  const path = url.pathname.replace(/^\/api\/(pbi|fabric)/, "");
  const response = await fetch(`${upstream.base}${path}${url.search}`, {
    method: c.req.method,
    headers: { Authorization: `Bearer ${token.token}`, "Content-Type": c.req.header("content-type") ?? "application/json" },
    body: ["GET", "HEAD"].includes(c.req.method) ? undefined : await c.req.arrayBuffer(),
    redirect: "follow",
  });
  const headers = new Headers();
  for (const name of ["content-type", "content-disposition", "location", "retry-after", "x-ms-operation-id"]) {
    const value = response.headers.get(name);
    if (value) headers.set(name, value);
  }
  return new Response(response.body, { status: response.status, headers });
});
