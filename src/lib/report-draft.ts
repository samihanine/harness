/**
 * A virtual copy of a report's definition (PBIR) that the AI reads and changes with JSON Patch
 * operations, saved in one go when it has finished. Saving replays the logged operations on a fresh
 * copy of the definition (read after saving the editor), so changes made meanwhile by the user or by
 * other tools are kept; the previous definition is kept to undo the last save.
 */
import type { Ref } from "./pbi";
import { definition, partText, textPayload, writeDefinition, type Part } from "./report-def";

export type PatchOp = { op: "add" | "replace" | "remove"; path: string; value?: unknown };
type Draft = { files: Map<string, unknown>; log: { file: string; ops: PatchOp[] }[] };

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
function applyOp(doc: unknown, { op, path, value }: PatchOp) {
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
function check(file: string, json: unknown) {
  const o = json as Record<string, any>;
  if (!o || typeof o !== "object" || Array.isArray(o)) throw new Error(`${file} must stay a JSON object.`);
  if (file.endsWith("/visual.json")) {
    if (typeof o.name !== "string") throw new Error(`${file}: "name" is required.`);
    const p = o.position;
    if (!p || !["x", "y", "width", "height"].every((k) => typeof p[k] === "number")) throw new Error(`${file}: "position" needs numbers x, y, width, height.`);
    if (!o.visual?.visualType && !o.visualGroup) throw new Error(`${file}: "visual.visualType" is required.`);
  }
  if (file.endsWith("/page.json") && (typeof o.name !== "string" || typeof o.width !== "number" || typeof o.height !== "number")) throw new Error(`${file}: "name", "width" and "height" are required.`);
}

/* ---------------------------------- Draft ---------------------------------- */

const pbirError = (e: unknown) =>
  new Error(`${e instanceof Error ? e.message : e} — the JSON editing needs the report in PBIR format with edit rights; use format_visual and the other tools instead.`);

/** The report's draft, made on first use: the editor is saved first (`save`) so the definition is current. */
async function open(report: Ref, save: () => Promise<unknown>) {
  let draft = drafts.get(report.id);
  if (draft) return draft;
  await save().catch(() => undefined);
  const parts = await definition(report).catch((e) => {
    throw pbirError(e);
  });
  const files = new Map<string, unknown>();
  for (const p of parts) if (p.path.endsWith(".json")) files.set(p.path, JSON.parse(partText(p)));
  drafts.set(report.id, (draft = { files, log: [] }));
  return draft;
}

/** A file of the draft (its current virtual state), as JSON text; `paths` lists the files when it is missing. */
export async function readDraftFile(report: Ref, save: () => Promise<unknown>, file: string) {
  const draft = await open(report, save);
  const json = draft.files.get(file);
  if (json === undefined) {
    const near = [...draft.files.keys()].filter((f) => f.includes(file.split("/").slice(-2).join("/")) || f.endsWith(file)).slice(0, 8);
    throw new Error(`No file ${file} in the report definition.${near.length ? ` Close: ${near.join(", ")}` : ""}`);
  }
  // "$schema" says nothing useful to the AI: left out of the reading (kept in the file).
  const { $schema: _, ...shown } = json as Record<string, unknown>;
  return JSON.stringify(shown, null, 1);
}

/** Applies operations to a file of the draft (all or none), checks it, and logs them for the save. */
export async function patchDraftFile(report: Ref, save: () => Promise<unknown>, file: string, ops: PatchOp[]) {
  const draft = await open(report, save);
  const current = draft.files.get(file);
  if (current === undefined) throw new Error(`No file ${file} in the report definition: read it first.`);
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
  draft.log.push({ file, ops });
  return `${ops.length} change(s) to ${file} (in the draft; saved at the end of your answer).`;
}

/** Files changed in the draft and not saved yet. */
export const pendingFiles = (report: Ref) => [...new Set(drafts.get(report.id)?.log.map((l) => l.file) ?? [])];

export const discardDraft = (report: Ref) => void drafts.delete(report.id);

/**
 * Saves the draft: the editor is saved, a fresh definition read, the logged operations replayed on it,
 * the result checked and written; the draft is then dropped (kept on failure, to retry). Returns a
 * summary, or undefined when there was nothing to save. The caller reloads the editor (see rewrite).
 */
export async function saveDraft(report: Ref, save: () => Promise<unknown>) {
  const draft = drafts.get(report.id);
  if (!draft?.log.length) return undefined;
  await save().catch(() => undefined);
  const parts = await definition(report).catch((e) => {
    throw pbirError(e);
  });
  const changed = new Map<string, unknown>();
  for (const { file, ops } of draft.log) {
    const part = parts.find((p) => p.path === file);
    if (!part && !changed.has(file)) throw new Error(`${file} is not in the report anymore (deleted meanwhile?): the changes to it cannot be saved.`);
    const json = changed.get(file) ?? JSON.parse(partText(part!));
    for (const op of ops) applyOp(json, op);
    check(file, json);
    changed.set(file, json);
  }
  const next = parts.map((p) => (changed.has(p.path) ? { ...p, payload: textPayload(JSON.stringify(changed.get(p.path), null, 2)) } : p));
  await writeDefinition(report, next);
  previous.set(report.id, parts);
  drafts.delete(report.id);
  return `Saved: ${changed.size} file(s) of the report definition changed (${[...changed.keys()].map((f) => f.replace(/^definition\//, "")).join(", ")}).`;
}

/** Writes back the definition as it was before the last save of the draft. */
export async function undoLastSave(report: Ref) {
  const parts = previous.get(report.id);
  if (!parts) throw new Error("No saved JSON change to undo in this session.");
  await writeDefinition(report, parts);
  previous.delete(report.id);
  drafts.delete(report.id);
  return "The report definition is back to how it was before the last JSON changes.";
}
