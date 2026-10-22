/** Stand-in for src/server/ms.ts: the eval never calls Microsoft (Power BI, SharePoint, Forms). */
const off = () => {
  throw new Error("Microsoft calls are disabled in the agent eval.");
};
const fn = Object.assign(off, { handler: off });
export const authStatus = fn, signIn = fn, signOut = fn, getToken = fn, ms = fn, refreshModel = fn, uploadFile = fn, exportPbix = fn, downloadFile = fn, readForm = fn, formsApi = fn, readItem = fn, writeItem = fn;
export type Service = string;
export type MsRequest = Record<string, unknown>;
