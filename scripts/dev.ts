/** `bun dev`: the server (API, port 3301) and Vite (the app, port 3300, /api forwarded); stopping one stops both. */
export {};
const procs = [
  Bun.spawn(["bun", "--watch", "server/main.ts"], { stdio: ["inherit", "inherit", "inherit"] }),
  Bun.spawn(["bunx", "--bun", "vite", "dev", "--port", "3300", "--strictPort"], { stdio: ["inherit", "inherit", "inherit"] }),
];
const stopAll = () => procs.forEach((p) => p.kill());
process.on("SIGINT", stopAll);
process.on("SIGTERM", stopAll);
await Promise.race(procs.map((p) => p.exited));
stopAll();
