/** Images stored as files in a local folder; cells hold the file name. */
const urls = new Map<string, Promise<string | null>>();

export function imageUrl(folder: FileSystemDirectoryHandle | undefined, name: unknown) {
  if (!folder || typeof name !== "string" || !name) return Promise.resolve(null);
  const key = `${folder.name}/${name}`;
  if (!urls.has(key))
    urls.set(
      key,
      folder
        .getFileHandle(name)
        .then((handle) => handle.getFile())
        .then((file) => URL.createObjectURL(file))
        .catch(() => null),
    );
  return urls.get(key)!;
}

/** Copies a file into the folder (unique name) and returns its name. */
export async function saveImage(folder: FileSystemDirectoryHandle, file: File) {
  const dot = file.name.lastIndexOf(".");
  const base = (dot > 0 ? file.name.slice(0, dot) : file.name).replace(/[^\w-]+/g, "-").slice(0, 40) || "image";
  const ext = dot > 0 ? file.name.slice(dot) : ".png";
  let name = `${base}${ext}`;
  for (let i = 2; await exists(folder, name); i++) name = `${base}-${i}${ext}`;
  const handle = await folder.getFileHandle(name, { create: true });
  const stream = await handle.createWritable();
  await stream.write(file);
  await stream.close();
  return name;
}

const exists = (folder: FileSystemDirectoryHandle, name: string) =>
  folder.getFileHandle(name).then(
    () => true,
    () => false,
  );

/** Image files of the folder (for picking an existing one). */
export async function listImages(folder: FileSystemDirectoryHandle) {
  const names: string[] = [];
  for await (const [name, handle] of folder.entries())
    if (handle.kind === "file" && /\.(png|jpe?g|gif|webp|svg|avif)$/i.test(name)) names.push(name);
  return names.sort();
}
