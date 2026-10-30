/**
 * The app's options in a JSON file on SharePoint / OneDrive (no database): read at start when a link
 * is set in this browser (Settings › Config), written by an admin from the same page.
 */
import { applyConfig, currentConfig, type HubConfig } from "@/config";
import { shareId } from "./excel";
import { graph } from "./ms";
import { settings } from "@/lib/store";

const FILE = "hub-config.json";

/** Reads the file of the saved link (if any) and applies it; returns what was read. */
export async function loadHubConfig(link?: string) {
  const url = link ?? (await settings.get()).configLink;
  if (!url) return undefined;
  const config = (await graph(`/shares/${shareId(url)}/driveItem/content`)) as Partial<HubConfig>;
  if (!config || typeof config !== "object") throw new Error("This file is not a JSON configuration.");
  applyConfig(config);
  return config;
}

/**
 * Writes the current options: into the linked file, or into `hub-config.json` of the linked folder
 * (created). Returns the file's link, kept as the one to read at start.
 */
export async function saveHubConfig(link: string) {
  const item = await graph(`/shares/${shareId(link)}/driveItem?$select=id,name,folder,webUrl,parentReference`);
  const drive = item.parentReference.driveId as string;
  const path = item.folder ? `/drives/${drive}/items/${item.id}:/${FILE}:/content` : `/drives/${drive}/items/${item.id}/content`;
  const saved = await graph(path, "PUT", currentConfig());
  const fileLink = (saved?.webUrl as string | undefined) ?? link;
  await settings.set({ configLink: fileLink });
  return fileLink;
}
