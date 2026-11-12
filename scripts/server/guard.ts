/**
 * Who may call the local server. It holds the user's Microsoft session, so it only answers the app:
 * it listens on the loopback interface, checks the Host (a site whose name is pointed at 127.0.0.1
 * — DNS rebinding — is refused), the Origin (only the app's own addresses; no "*"), and a custom
 * header (a page of another site cannot send it without a CORS preflight, which is refused).
 */
export const PORT = Number(process.env.TOKEN_PORT ?? 3301);
/** The dev server's port (bun dev: APP_PORT, default 3300). */
const APP_PORT = Number(process.env.APP_PORT ?? 3300);
export const HOST = process.env.TOKEN_HOST ?? "127.0.0.1";

const LOOPBACK = /^(localhost|127\.0\.0\.1|\[::1\])(:\d+)?$/i;
/** The app's addresses: the dev server (APP_PORT), this server (it serves the exported app), and HUB_ORIGINS (comma-separated: another port, another machine). */
const ORIGINS = new Set(
  [`http://localhost:${APP_PORT}`, `http://127.0.0.1:${APP_PORT}`, `http://localhost:${PORT}`, `http://127.0.0.1:${PORT}`, ...(process.env.HUB_ORIGINS ?? "").split(",")]
    .map((o) => o.trim().replace(/\/+$/, ""))
    .filter(Boolean),
);
/** The exported app opened as a file has the origin "null", which any sandboxed page also has: allowed only when asked. */
if (process.env.HUB_ALLOW_FILE === "1") ORIGINS.add("null");

export const CLIENT_HEADER = "x-hub-client";
const EXPOSED = "Location, Retry-After, ETag, Content-Disposition, Content-Type, Content-Length, x-ms-request-id";

/** CORS headers for an allowed origin (never "*"). */
export function cors(req: Request): Record<string, string> {
  const origin = req.headers.get("origin");
  if (!origin || !ORIGINS.has(origin)) return {};
  return {
    "Access-Control-Allow-Origin": origin,
    Vary: "Origin",
    "Access-Control-Allow-Methods": "GET, POST, PUT, PATCH, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": `Content-Type, If-Match, Prefer, Accept, ${CLIENT_HEADER}`,
    "Access-Control-Expose-Headers": EXPOSED,
    "Access-Control-Max-Age": "600",
  };
}

/** Why a request is refused (undefined: allowed). `api`: a call of the app (not a page it serves). */
export function refusal(req: Request, api = true): string | undefined {
  if (!LOOPBACK.test(req.headers.get("host") ?? "")) return "This server only answers on localhost.";
  const origin = req.headers.get("origin");
  if (origin) {
    if (!ORIGINS.has(origin)) return `Origin ${origin} is not the app (allowed: ${[...ORIGINS].join(", ")}; more with HUB_ORIGINS).`;
  } else {
    // No Origin: a same-origin request of a page served here, or a local program (scripts) — never another site's page.
    const site = req.headers.get("sec-fetch-site");
    if (site && site !== "same-origin" && site !== "none") return "Cross-site requests are refused.";
  }
  if (api && req.method !== "OPTIONS" && req.headers.get(CLIENT_HEADER) !== "1") return `Missing header ${CLIENT_HEADER}.`;
  return undefined;
}

export const json = (req: Request, body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff", ...cors(req) } });
