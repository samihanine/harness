/**
 * A virtual copy of a report's definition (PBIR) that the AI reads and changes — targeted JSON Patch
 * operations, new files (pages, visuals), deleted files — saved in one go when it has finished. Saving
 * replays the logged changes on a fresh copy of the definition (read after saving the editor), so
 * changes made meanwhile by the user are kept; the previous definition is kept to undo the last save.
 */
import type { Ref } from "./pbi";
import { definition, partText, textPayload, writeDefinition, type Part } from "./report-def";

export type PatchOp = { op: "add" | "replace" | "remove"; path: string; value?: unknown };
type Change = { kind: "patch"; file: string; ops: PatchOp[] } | { kind: "put"; file: string; json: unknown } | { kind: "delete"; prefix: string };
type Draft = { files: Map<string, unknown>; log: Change[] };

const drafts = new Map<string, Draft>();
const previous = new Map<string, Part[]>();

/* ------------------------------- JSON Pointer ------------------------------- */

const tokens = (pointer: string) => {
  if (pointer === "" || pointer === "/") return [];
  if (!pointer.startsWith("/")) throw new Error(`Path "${pointer}" must start with "/" (e.g. /visual/objects/title).`);
  return pointer.slice(1).split("/").map((t) => t.replace(/~1/g, "/").replace(/~0/g, "~"));
};

/**
 * One operation on a JSON value (in place). "add" also creates the missing objects on its way (the
 * usual case: a formatting object the visual does not have yet); "-" appends to an array.
 */
export function applyOp(doc: unknown, { op, path, value }: PatchOp) {
  const keys = tokens(path);
  if (!keys.length) throw new Error("The whole file cannot be replaced: target a key inside it.");
  let node = doc as Record<string, unknown> | unknown[];
  for (const [i, key] of keys.slice(0, -1).entries()) {
    let next = Array.isArray(node) ? node[Number(key)] : (node as Record<string, unknown>)[key];
    if (next === undefined || next === null) {
      if (op !== "add") throw new Error(`${path}: "${keys.slice(0, i + 1).join("/")}" does not exist (read the file first, or use "add").`);
      next = /^\d+$/.test(keys[i + 1]) || keys[i + 1] === "-" ? [] : {};
      if (Array.isArray(node)) node[Number(key)] = next;
      else (node as Record<string, unknown>)[key] = next;
    }
    if (typeof next !== "object") throw new Error(`${path}: "${keys.slice(0, i + 1).join("/")}" is a value, not an object.`);
    node = next as Record<string, unknown> | unknown[];
  }
  const last = keys[keys.length - 1];
  if (Array.isArray(node)) {
    const index = last === "-" ? node.length : Number(last);
    if (Number.isNaN(index)) throw new Error(`${path}: "${last}" is not an array index.`);
    if (op === "remove") node.splice(index, 1);
    else if (op === "add") node.splice(index, 0, value);
    else node[index] = value;
  } else {
    if (op !== "add" && !(last in node)) throw new Error(`${path}: "${last}" does not exist${op === "replace" ? ' (use "add")' : ""}.`);
    if (op === "remove") delete node[last];
    else node[last] = value;
  }
}

/** The structure Power BI requires of a file (the service checks the rest when saving). */
export function check(file: string, json: unknown) {
  const o = json as Record<string, any>;
  if (!o || typeof o !== "object" || Array.isArray(o)) throw new Error(`${file} must stay a JSON object.`);
  if (file.endsWith("/visual.json")) {
    if (typeof o.name !== "string") throw new Error(`${file}: "name" is required.`);
    const p = o.position;
    if (!p || !["x", "y", "width", "height"].every((k) => typeof p[k] === "number")) throw new Error(`${file}: "position" needs numbers x, y, width, height.`);
    if (!o.visual?.visualType && !o.visualGroup) throw new Error(`${file}: "visual.visualType" is required.`);
    // The shapes Power BI checks on save, checked here so a wrong patch is refused at once (not at the end).
    for (const group of ["objects", "visualContainerObjects"]) {
      const objects = o.visual?.[group];
      if (objects === undefined) continue;
      if (!objects || typeof objects !== "object" || Array.isArray(objects)) throw new Error(`${file}: "visual.${group}" must be an object {<object name>: [{"properties": {…}}]}.`);
      for (const [name, list] of Object.entries(objects)) {
        if (!Array.isArray(list)) throw new Error(`${file}: "visual.${group}.${name}" must be a list [{"properties": {…}}]${list && typeof list === "object" ? ` — "${name}" is not an object name (write visual.${group}.<object>, e.g. visual.objects.image)` : ""}.`);
        if (list.some((e) => !e || typeof e !== "object" || typeof (e as any).properties !== "object")) throw new Error(`${file}: each entry of "visual.${group}.${name}" needs "properties": {…}.`);
      }
    }
    const state = o.visual?.query?.queryState;
    if (state !== undefined && Object.entries(state ?? {}).some(([, r]: [string, any]) => !Array.isArray(r?.projections))) throw new Error(`${file}: each role of "visual.query.queryState" needs "projections": [...] (use set_visual_fields).`);
  }
  if (file.endsWith("/page.json") && (typeof o.name !== "string" || typeof o.width !== "number" || typeof o.height !== "number")) throw new Error(`${file}: "name", "width" and "height" are required.`);
  // Filters (report, page, visual): the shape Power BI checks on save.
  for (const f of o.filterConfig?.filters ?? []) {
    if (!f || typeof f.name !== "string" || !f.field) throw new Error(`${file}: each filter of "filterConfig.filters" needs "name" and "field".`);
    if (f.filter !== undefined && (!Array.isArray(f.filter?.From) || !Array.isArray(f.filter?.Where)))
      throw new Error(`${file}: filter "${f.name}": "filter" must be {"Version": 2, "From": [...], "Where": [...]}. For a Top N, use create_visual "top" or set_visual_top_n (they write it for you).`);
  }
  if (file.endsWith("pages/pages.json") && !Array.isArray(o.pageOrder)) throw new Error(`${file}: "pageOrder" must be a list.`);
}

/* ---------------------------------- Draft ---------------------------------- */

const pbirError = (e: unknown) => new Error(`${e instanceof Error ? e.message : e} — editing needs the report in PBIR format with edit rights.`);
const isJson = (path: string) => path.endsWith(".json") || /RegisteredResources\/[^/.]+$/.test(path);

/** The report's draft, made on first use: the editor is saved first (`save`) so the definition is current. */
async function open(report: Ref, save?: () => Promise<unknown>) {
  let draft = drafts.get(report.id);
  if (draft) return draft;
  await save?.().catch(() => undefined);
  const parts = await definition(report).catch((e) => {
    throw pbirError(e);
  });
  const files = new Map<string, unknown>();
  for (const p of parts)
    if (isJson(p.path))
      try {
        files.set(p.path, JSON.parse(partText(p)));
      } catch {
        // not JSON after all (e.g. an image without extension): left out of the draft
      }
  drafts.set(report.id, (draft = { files, log: [] }));
  return draft;
}

/** Every JSON file of the draft (its current virtual state); `save` saves the editor before the first read. */
export async function draftFiles(report: Ref, save?: () => Promise<unknown>) {
  return (await open(report, save)).files;
}

/** A file of the draft as JSON (a copy); undefined when missing. */
export async function readDraftJson<T = any>(report: Ref, save: () => Promise<unknown>, file: string): Promise<T | undefined> {
  const json = (await open(report, save)).files.get(file);
  return json === undefined ? undefined : structuredClone(json as T);
}

/** A file of the draft as text for the AI ("$schema" left out: it says nothing useful). */
export async function readDraftFile(report: Ref, save: () => Promise<unknown>, file: string) {
  const draft = await open(report, save);
  const json = draft.files.get(file);
  if (json === undefined) {
    const near = [...draft.files.keys()].filter((f) => f.includes(file.split("/").slice(-2).join("/")) || f.endsWith(file)).slice(0, 8);
    throw new Error(`No file ${file} in the report definition.${near.length ? ` Close: ${near.join(", ")}` : ""}`);
  }
  const { $schema: _, ...shown } = json as Record<string, unknown>;
  return JSON.stringify(shown, null, 1);
}

/** Applies operations to a file of the draft (all or none), checks it, and logs them for the save. */
export async function patchDraftFile(report: Ref, save: () => Promise<unknown>, file: string, ops: PatchOp[]) {
  const draft = await open(report, save);
  const current = draft.files.get(file);
  if (current === undefined) throw new Error(`No file ${file} in the report definition.`);
  const next = structuredClone(current);
  ops.forEach((op, i) => {
    try {
      applyOp(next, op);
    } catch (e) {
      throw new Error(`Operation ${i + 1} (${op.op} ${op.path}): ${String(e instanceof Error ? e.message : e).replace(/\.$/, "")}. Nothing was changed.`);
    }
  });
  check(file, next);
  draft.files.set(file, next);
  draft.log.push({ kind: "patch", file, ops });
  return `${ops.length} change(s) to ${file.replace(/^definition\//, "")} (in the draft; saved at the end of your answer).`;
}

/** Writes a whole file in the draft (a new page or visual, or a file rebuilt by the app). */
export async function putDraftFile(report: Ref, save: () => Promise<unknown>, file: string, json: unknown) {
  check(file, json);
  const draft = await open(report, save);
  draft.files.set(file, structuredClone(json));
  draft.log.push({ kind: "put", file, json: structuredClone(json) });
}

/** Deletes the files under a path (a page folder, a visual folder) in the draft. */
export async function deleteDraftFiles(report: Ref, save: () => Promise<unknown>, prefix: string) {
  const draft = await open(report, save);
  const gone = [...draft.files.keys()].filter((f) => f.startsWith(prefix));
  if (!gone.length) throw new Error(`Nothing under ${prefix} in the report definition.`);
  gone.forEach((f) => draft.files.delete(f));
  draft.log.push({ kind: "delete", prefix });
  return gone.length;
}

/** Files changed in the draft and not saved yet. */
export const pendingFiles = (report: Ref) => [...new Set(drafts.get(report.id)?.log.map((c) => (c.kind === "delete" ? `${c.prefix} (deleted)` : c.file)) ?? [])];
export const hasDraft = (report: Ref) => drafts.has(report.id);
export const discardDraft = (report: Ref) => void drafts.delete(report.id);

/**
 * Saves the draft: the editor is saved, a fresh definition read, the logged changes replayed on it, the
 * result checked and written; the draft is then dropped (kept on failure, to retry). Returns a summary,
 * or undefined when there was nothing to save. The caller reloads the editor.
 */
export async function saveDraft(report: Ref, save: () => Promise<unknown>) {
  const draft = drafts.get(report.id);
  if (!draft?.log.length) return undefined;
  await save().catch(() => undefined);
  const parts = await definition(report).catch((e) => {
    throw pbirError(e);
  });
  const files = new Map<string, unknown>();
  const read = (file: string) => {
    if (files.has(file)) return files.get(file);
    const part = parts.find((p) => p.path === file);
    if (!part) throw new Error(`${file} is not in the report anymore (deleted meanwhile?): the changes to it cannot be saved.`);
    return JSON.parse(partText(part));
  };
  const deleted = new Set<string>();
  for (const change of draft.log) {
    if (change.kind === "delete") {
      for (const p of [...parts.map((x) => x.path), ...files.keys()]) if (p.startsWith(change.prefix)) (deleted.add(p), files.delete(p));
      continue;
    }
    if (change.kind === "put") {
      files.set(change.file, structuredClone(change.json));
      deleted.delete(change.file);
      continue;
    }
    const json = read(change.file);
    for (const op of change.ops) applyOp(json, op);
    files.set(change.file, json);
  }
  for (const [file, json] of files) check(file, json);
  const payload = (json: unknown) => textPayload(JSON.stringify(json, null, 2));
  const next: Part[] = [
    ...parts.filter((p) => !deleted.has(p.path)).map((p) => (files.has(p.path) ? { ...p, payload: payload(files.get(p.path)) } : p)),
    ...[...files.keys()].filter((f) => !parts.some((p) => p.path === f)).map((f) => ({ path: f, payload: payload(files.get(f)), payloadType: "InlineBase64" })),
  ];
  await writeDefinition(report, next);
  previous.set(report.id, parts);
  drafts.delete(report.id);
  const changed = [...files.keys()].map((f) => f.replace(/^definition\//, ""));
  return `Saved: ${changed.length} file(s) written${deleted.size ? `, ${deleted.size} deleted` : ""} (${[...changed, ...[...deleted].map((d) => `−${d.replace(/^definition\//, "")}`)].slice(0, 8).join(", ")}).`;
}

/** Writes back the definition as it was before the last save of the draft. */
export async function undoLastSave(report: Ref) {
  const parts = previous.get(report.id);
  if (!parts) throw new Error("No saved change to undo in this session.");
  await writeDefinition(report, parts);
  previous.delete(report.id);
  drafts.delete(report.id);
  return "The report is back to how it was before the last saved changes.";
}
