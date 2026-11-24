import { graph } from "@server/ms/api";
const d = await graph("/me/drive").catch((e) => ({ error: String(e) }));
console.log(JSON.stringify(d).slice(0, 300));
