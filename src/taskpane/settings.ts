/* global Office */
import { MAX_RECENTS, RECENTS_KEY } from "./config";
import type { DriveFolder } from "./graph";

export function getRecentFolders(): DriveFolder[] {
  const value = Office.context.roamingSettings.get(RECENTS_KEY);
  return Array.isArray(value) ? (value as DriveFolder[]).slice(0, MAX_RECENTS) : [];
}

export async function rememberFolder(folder: DriveFolder): Promise<void> {
  const existing = getRecentFolders().filter((item) => item.id !== folder.id);
  Office.context.roamingSettings.set(RECENTS_KEY, [folder, ...existing].slice(0, MAX_RECENTS));
  await new Promise<void>((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "Einstellungen konnten nicht gespeichert werden."));
    });
  });
}
