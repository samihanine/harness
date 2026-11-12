/**
 * `bun dev`: the app (Vite, port 3300, or APP_PORT) and the local server (port 3301) together; stopping one stops both.
 * With another APP_PORT the server is told that origin is the app (scripts/server/guard.ts).
 * The server checks TLS certificates against the system's store (--use-system-ca: a company proxy's own
 * certificate authority is trusted the way the browser trusts it); verification is never turned off.
 */
const APP_PORT = String(process.env.APP_PORT ?? 3300);
const procs = [
  Bun.spawn(["bun", "--use-system-ca", "scripts/token-server.ts"], { stdio: ["inherit", "inherit", "inherit"], env: { ...process.env, APP_PORT } }),
  Bun.spawn(["bunx", "--bun", "vite", "dev", "--port", APP_PORT], { stdio: ["inherit", "inherit", "inherit"], env: { ...process.env, APP_PORT } }),
];
const stop = () => procs.forEach((p) => p.kill());
process.on("SIGINT", stop);
process.on("SIGTERM", stop);
await Promise.race(procs.map((p) => p.exited));
stop();
