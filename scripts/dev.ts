/** `bun dev`: the app (Vite, port 3300) and the local token server (port 3301) together; stopping one stops both. */
const procs = [
  Bun.spawn(["bun", "scripts/token-server.ts"], { stdio: ["inherit", "inherit", "inherit"] }),
  Bun.spawn(["bunx", "--bun", "vite", "dev", "--port", "3300"], { stdio: ["inherit", "inherit", "inherit"] }),
];
const stop = () => procs.forEach((p) => p.kill());
process.on("SIGINT", stop);
process.on("SIGTERM", stop);
await Promise.race(procs.map((p) => p.exited));
stop();
