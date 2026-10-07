/** Builds the backends of a file link (asking for local access or a Microsoft sign-in when needed). */
import { handles, permission } from "@/lib/store";
import type { FileLink } from "@/lib/store";
import { SignInRequiredError, getToken } from "./graph";
import { DriveImages, LocalImages } from "./images";
import { LocalTable } from "./local";
import { SharePointTable } from "./sharepoint";
import type { ImageStore, TableBackend } from "./types";

export type Opened = { backend: TableBackend; images?: ImageStore };

/** Thrown when the browser needs a click to allow access to local files. */
export class PermissionNeededError extends Error {
  constructor(readonly fileName: string) {
    super(`Allow access to ${fileName}`);
  }
}

export async function openStorage(link: FileLink, ask: boolean): Promise<Opened> {
  const needsGraph = link.source.kind === "sharepoint" || link.images?.kind === "sharepoint";
  if (needsGraph) await getToken(); // throws SignInRequiredError

  let backend: TableBackend;
  if (link.source.kind === "sharepoint") backend = new SharePointTable(link.source);
  else {
    const file = await handles.file(link.id);
    if (!file) throw new Error("The file handle is missing: add the file again.");
    if (!(await permission(file, ask))) throw new PermissionNeededError(link.source.name);
    backend = new LocalTable(file);
  }

  let images: ImageStore | undefined;
  if (link.images?.kind === "sharepoint") images = new DriveImages(link.images);
  else if (link.images?.kind === "local") {
    const folder = await handles.images(link.id);
    if (folder) {
      if (!(await permission(folder, ask))) throw new PermissionNeededError(folder.name);
      images = new LocalImages(folder);
    }
  }
  return { backend, images };
}

export { SignInRequiredError };
