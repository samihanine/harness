/**
 * `bun dev`: the app (Vite, port 3300) and the local server (port 3301) together; stopping one stops both.
 * The server checks TLS certificates against the system's store (--use-system-ca: a company proxy's own
 * certificate authority is trusted the way the browser trusts it); verification is never turned off.
 */
const procs = [
  Bun.spawn(["bun", "--use-system-ca", "scripts/token-server.ts"], { stdio: ["inherit", "inherit", "inherit"] }),
  Bun.spawn(["bunx", "--bun", "vite", "dev", "--port", "3300"], { stdio: ["inherit", "inherit", "inherit"] }),
];
const stop = () => procs.forEach((p) => p.kill());
process.on("SIGINT", stop);
process.on("SIGTERM", stop);
await Promise.race(procs.map((p) => p.exited));
stop();
