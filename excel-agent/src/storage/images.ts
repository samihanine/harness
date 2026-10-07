/** Image stores: a local folder (cells hold file names) or a SharePoint / OneDrive folder (cells hold links). */
import { graph, isDriveLink } from "./graph";
import type { DriveRef, ImageStore } from "./types";

const IMAGE = /\.(png|jpe?g|gif|webp|svg|avif|bmp)$/i;
const urls = new Map<string, Promise<string | null>>();
const cached = (key: string, load: () => Promise<string | null>) => {
  if (!urls.has(key)) urls.set(key, load().catch(() => null));
  return urls.get(key)!;
};

const safeName = (file: File) => {
  const dot = file.name.lastIndexOf(".");
  const base = (dot > 0 ? file.name.slice(0, dot) : file.name).replace(/[^\w-]+/g, "-").slice(0, 40) || "image";
  return { base, ext: dot > 0 ? file.name.slice(dot) : ".png" };
};

export class LocalImages implements ImageStore {
  constructor(private folder: FileSystemDirectoryHandle) {}
  get label() {
    return this.folder.name;
  }

  async save(file: File) {
    const { base, ext } = safeName(file);
    let name = `${base}${ext}`;
    for (let i = 2; await this.folder.getFileHandle(name).then(() => true, () => false); i++) name = `${base}-${i}${ext}`;
    const stream = await (await this.folder.getFileHandle(name, { create: true })).createWritable();
    await stream.write(file);
    await stream.close();
    return name;
  }

  async list() {
    const names: { id: string; name: string }[] = [];
    for await (const [name, handle] of this.folder.entries()) if (handle.kind === "file" && IMAGE.test(name)) names.push({ id: name, name });
    return names.sort((a, b) => a.name.localeCompare(b.name));
  }

  async pick(id: string) {
    return id;
  }

  resolve(value: string) {
    return cached(`${this.folder.name}/${value}`, async () => URL.createObjectURL(await (await this.folder.getFileHandle(value)).getFile()));
  }
}

/** Images uploaded to a drive folder; the cell keeps a sharing link (readable outside the app too). */
export class DriveImages implements ImageStore {
  constructor(private folder: DriveRef) {}
  get label() {
    return this.folder.name;
  }

  async save(file: File) {
    if (file.size > 4 * 1024 * 1024) throw new Error("Images up to 4 MB.");
    const { base, ext } = safeName(file);
    const item = await graph<{ id: string }>(
      `/drives/${this.folder.driveId}/items/${this.folder.itemId}:/${encodeURIComponent(base + ext)}:/content?@microsoft.graph.conflictBehavior=rename`,
      { method: "PUT", body: file, headers: { "Content-Type": file.type || "application/octet-stream" } },
    );
    return this.pick(item.id);
  }

  async list() {
    const { value } = await graph<{ value: { id: string; name: string; file?: unknown }[] }>(
      `/drives/${this.folder.driveId}/items/${this.folder.itemId}/children?$select=id,name,file&$top=500`,
    );
    return value.filter((item) => item.file && IMAGE.test(item.name)).map(({ id, name }) => ({ id, name }));
  }

  /** Sharing link of an image: anonymous when the tenant allows it, else for the organization. */
  async pick(id: string) {
    const link = (scope: string) =>
      graph<{ link: { webUrl: string } }>(`/drives/${this.folder.driveId}/items/${id}/createLink`, { method: "POST", json: { type: "view", scope } });
    return (await link("anonymous").catch(() => link("organization"))).link.webUrl;
  }

  resolve(value: string) {
    return resolveImage(value);
  }
}

/**
 * Displayable URL of any image value: drive links are read with the user's token,
 * other web links are used as they are. Local file names need the LocalImages store.
 */
export function resolveImage(value: string): Promise<string | null> {
  if (isDriveLink(value)) {
    const encoded = `u!${btoa(value).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-")}`;
    return cached(value, async () => URL.createObjectURL(await graph<Blob>(`/shares/${encoded}/driveItem/content`)));
  }
  return Promise.resolve(/^https?:\/\//i.test(value) ? value : null);
}
