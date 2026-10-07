import { useState } from "react";
import { FolderIcon, XIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import type { DriveRef } from "@/storage/types";
import { DriveLinkField, MicrosoftSignIn, StorageTabs, useMicrosoft } from "./storage-ui";
import type { StorageKind } from "./storage-ui";

/** Chosen images folder: a local folder handle, or a SharePoint / OneDrive folder. */
export type ImagesChoice = { kind: "local"; handle: FileSystemDirectoryHandle } | ({ kind: "sharepoint" } & DriveRef);

export function ImagesPicker({ value, onChange, initialKind = "local" }: { value?: ImagesChoice; onChange: (choice?: ImagesChoice) => void; initialKind?: StorageKind }) {
  const [kind, setKind] = useState<StorageKind>(value?.kind ?? initialKind);
  const microsoft = useMicrosoft();
  if (value?.kind === "local")
    return (
      <div className="flex h-8 items-center gap-2 rounded-lg border bg-muted/40 px-2.5 [&_svg]:size-3.5 [&_svg]:text-muted-foreground">
        <FolderIcon />
        <span className="flex-1 truncate">{value.handle.name}</span>
        <button type="button" aria-label="Clear" onClick={() => onChange(undefined)}>
          <XIcon />
        </button>
      </div>
    );
  return (
    <div className="flex flex-col gap-1.5">
      <StorageTabs value={kind} onChange={setKind} />
      {kind === "local" ? (
        <Button
          variant="outline"
          size="sm"
          className="w-fit"
          disabled={!("showDirectoryPicker" in window)}
          onClick={() =>
            void window
              .showDirectoryPicker({ mode: "readwrite", id: "images" })
              .then((handle) => onChange({ kind: "local", handle }))
              .catch(() => undefined)
          }
        >
          <FolderIcon /> Choose a folder…
        </Button>
      ) : (
        <>
          <MicrosoftSignIn onSignedIn={() => void microsoft.refresh()} />
          {microsoft.status?.signedIn && (
            <DriveLinkField
              kind="folder"
              value={value?.kind === "sharepoint" ? value : undefined}
              onChange={(ref) => onChange(ref ? { kind: "sharepoint", ...ref } : undefined)}
              placeholder="Link to the images folder"
            />
          )}
        </>
      )}
    </div>
  );
}
