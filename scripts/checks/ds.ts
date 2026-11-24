import { pbi } from "@server/ms/api";
console.log(JSON.stringify((await pbi("/datasets")).value.map((d: any) => [d.id, d.name, d.targetStorageMode, d.configuredBy])));
