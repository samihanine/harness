# Security

How the app keeps the user's accounts and data safe, what each part is trusted with, and what to
check before changing any of it. The code named here is the reference; keep this file in step with it.

## The two parts

| Part | Runs | Holds | Never holds |
| --- | --- | --- | --- |
| The page (`src/`) | in the browser, served from `http://localhost:3300` (or by the local server) | the AI key (IndexedDB), the Power BI token (memory only), the library and conversations (IndexedDB) | the Microsoft refresh token, the Fabric and Graph tokens, the Confluence key |
| The local server (`scripts/token-server.ts`, `scripts/server/`) | on the user's computer, `127.0.0.1:3301` | the Microsoft session (refresh token), the Confluence settings and key, the preview engine | the AI key; any business logic |

The AI API is called by the page only (`src/agent/llm.ts`): the local server never sees prompts,
answers or the AI key. Everything else that needs a secret goes through the local server.

## Microsoft sign-in and tokens

- One device-code sign-in with the Microsoft Office public client: no app registration, no client
  secret, no admin consent. The refresh token is the only long-lived secret.
- **Where it is kept**: the operating system's credential store (Windows Credential Manager, macOS
  Keychain, libsecret) through `Bun.secrets`; where there is none, `.local/secrets.json`, readable by
  the user only (mode 600) and ignored by git (`scripts/server/secrets.ts`; `HUB_SECRETS=file` forces the file). The file of earlier
  versions (`.local/tokens.json`, in clear) is moved there at the first start, then deleted.
- **Calls to Microsoft go through the server**: the page asks `/ms/<service>/…`, the server adds the
  token and forwards to Power BI REST, Fabric or Graph (`microsoft()` in `scripts/token-server.ts`).
  The Fabric and Graph tokens never reach the browser. An absolute URL (an operation or download link
  a service gave back) is forwarded only to that service's own hosts, over https: the server is not
  an open proxy.
- **The one token in the page**: the Power BI token, because the embedding SDK draws reports with it
  in the browser. It is asked when needed, kept in memory (never localStorage / IndexedDB), and gone
  when the tab closes. Tokens stored in localStorage by earlier versions are removed at start
  (`src/services/ms.ts`).
- Sign-out (`POST /sign-out`) deletes the session from the store and the caches.

## Who may call the local server (`scripts/server/guard.ts`)

The server holds the user's session, so it answers the app only:

1. **Loopback only**: it listens on `127.0.0.1` (`TOKEN_HOST` to change), never on the network.
2. **Host check**: the `Host` header must be `localhost` / `127.0.0.1` — a site whose name is pointed
   at 127.0.0.1 (DNS rebinding) is refused.
3. **Origin allowlist**: `http://localhost:3300` (the dev app: `APP_PORT`), `http://127.0.0.1:3300`, and the server's own
   address (it serves the exported app). More with `HUB_ORIGINS=a,b`. CORS headers name that origin,
   never `*`. A request without `Origin` is accepted only when the browser says it is same-origin
   (`Sec-Fetch-Site`) or when it comes from a local program (scripts).
4. **A custom header** (`X-Hub-Client: 1`) on every call: another site's page cannot send it without
   a CORS preflight, which is refused for its origin.
5. The exported single file opened as `file://` has the origin `null`, which any sandboxed page
   also has: refused by default. Open the export through the server instead
   (`http://localhost:3301`, it serves `dist/hub.html`), or allow it knowingly with `HUB_ALLOW_FILE=1`.

## Confluence

The address, email, key and folder are kept by the local server, in the same credential store
(`/confluence-settings`): the page reads back everything but the key (only "a key is saved"). Calls
go through `/confluence/…`, where the server adds the key. A key saved in the browser by an earlier
version is handed to the server once, then removed from IndexedDB (`src/services/confluence.ts`).

## The AI key

Kept in this browser's IndexedDB (Settings › AI) and sent only to the AI API by the page: the
requirement is that AI calls stay in the front end. It is the one secret a script injected into the
page could read; the measures below are what stands in the way.

## The page

- **Content security policy** (written into `index.html` by `vite.config.ts`): scripts and styles
  from the app only, network calls only to the local server and the AI API's origin (read from
  `src/agent/llm.ts`), frames only from Power BI / Office, no plugins, no `<base>` change. A script
  slipped into the page could not load code from elsewhere nor send data to another site.
  `HUB_CSP_CONNECT` adds origins (a local server on another address); `HUB_NO_CSP=1` leaves the
  policy out (not recommended).
- **No HTML from data**: markdown (AI answers, guides, Confluence pages) is rendered by
  react-markdown without raw HTML; link and picture addresses are kept only when harmless
  (`safeUrl` in `src/lib/markdown/syntax.ts`: http(s), mailto, relative; `javascript:` and `data:`
  links are dropped). Links open in a new tab with `noopener noreferrer`. No `dangerouslySetInnerHTML`,
  no `eval`.
- **The AI's own JavaScript** (`run_js`) runs in a throwaway Web Worker without network access
  (`src/agent/sandbox.ts`), with a time limit.
- **Excel**: a cell starting like a formula (`=`, `+`, `-`, `@`) is written as text, so content
  written by the AI or pasted by a user never runs as a formula in someone's Excel (`toCell` in
  `src/features/excel/table.ts`). Files are written only if nobody saved them meanwhile (ETag).
- **Office files** (pptx, xlsx) are parsed as XML in the browser (`src/lib/ooxml`); XML written by the
  AI is parsed strictly and refused when invalid.
- The dev server listens on `localhost` only.

## The slides' preview

- **Microsoft Office (default)**: to draw a deck exactly as PowerPoint does, a work copy of it is put in
  the folder "Hub previews" of **the user's own OneDrive** (Graph, through the local server) and converted to
  PDF by Microsoft 365. The deck stays in the user's Microsoft 365 account (the same tenant, under their own
  rights); the copy is overwritten at each change and deleted when the deck is removed from the app. An
  account that cannot (no OneDrive, no right, a sensitivity label that blocks conversion) falls back to the
  local engine. Off with Settings › Power BI › Presenter (`PRESENTER_OFFICE`): nothing then leaves the computer.
- **Local engine**: LibreOffice compiled to WebAssembly, in a process of its own beside the local server
  (`scripts/server/office.ts`, `office-engine.mjs`). The deck goes from the page to the server on the same
  computer and nowhere else; nothing is written to disk. It reads the fonts installed on the computer that the
  deck names.
- The PDF is drawn by pdf.js in the page (worker inlined, `isEvalSupported: false`: no `eval`).

## TLS

The local server checks certificates against the system's store (`--use-system-ca`), so a company
proxy that inspects TLS with its own certificate authority works without turning verification off.
`NODE_TLS_REJECT_UNAUTHORIZED=0` is no longer set anywhere; do not bring it back.

## Dependencies and the repository

- `bun run audit`: known vulnerabilities of the dependencies and secrets committed by mistake;
  `bun run audit:fix`: the smallest safe versions (kept only if typecheck and build pass).
- `.local/`, `.env*` and `dist/` are ignored by git. Never commit a key, a token or a real tenant id.
- No telemetry: the app calls the local server, the AI API, and (through the server) Microsoft and
  Confluence — nothing else.

## Known limits

- The AI key lives in the browser (by requirement). Anyone with access to the browser profile can
  read it, as with any web app that stores a key client-side.
- The Power BI token is in the page's memory while a report is embedded (required by the SDK).
- The policy allows inline scripts (`'unsafe-inline'`): Vite's dev preamble and the single-file
  export need it. The network and frame rules are what limits an injected script.
- Without an OS credential store (some Linux setups), the session is in a user-only file, not encrypted.
- The local server trusts other programs running as the same user on the computer (they could read
  the same store): it protects against web pages and the network, not against local malware.

## Checklist for a change

- A new secret → the local server's store (`getSecret` / `setSecret`), never the browser, never a file in the repo.
- A new backend call → through the local server when it needs a secret; add its host to the allowlists, never a wildcard.
- A new origin the page must reach → `vite.config.ts` (CSP), with the reason.
- Text from the AI, a file or a remote page shown in the UI → through `Markdown` / React text, never as HTML.
- A new route of the local server → behind `refusal()`; state-changing routes are POST.
