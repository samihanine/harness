/** Stand-in for src/services/ms.ts: the eval never calls Microsoft (Power BI, Fabric, Graph). */
const off = (..._: unknown[]): never => {
  throw new Error("Microsoft calls are disabled in the agent eval (EVAL_LIVE=1 to use the real ones).");
};
export class SignInRequired extends Error {}
export const token = off, tokenStatus = off, signOut = off, powerBiToken = off, request = off, call = off, pbi = off, fabric = off, graph = off, refreshModel = off, exportPbix = off, readItem = off, writeItem = off, uploadFile = off;
export const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));
export type Service = "powerbi" | "fabric" | "graph";
