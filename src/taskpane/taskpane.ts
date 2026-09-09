/* global Office */
import { initializeAuth } from "./auth";
import { createFolder, getMessageMime, getOneDriveRoot, listFolders, uploadEml, type DriveFolder } from "./graph";
import { getCurrentMessage, makeDefaultFilename, sanitizeFilename, type CurrentMessage } from "./outlook";
import { getRecentFolders, rememberFolder } from "./settings";

let message: CurrentMessage;
let rootFolder: DriveFolder;
let selectedFolder: DriveFolder;
let browseFolder: DriveFolder;
let browseStack: DriveFolder[] = [];

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

function setStatus(text: string, kind: "info" | "success" | "error" = "info") {
  const el = $("status");
  el.textContent = text;
  el.dataset.kind = kind;
  el.hidden = !text;
}

function setBusy(busy: boolean) {
  $("saveButton").toggleAttribute("disabled", busy);
  $("chooseFolderButton").toggleAttribute("disabled", busy);
  document.body.dataset.busy = busy ? "true" : "false";
}

function renderMessage() {
  $("subject").textContent = message.subject;
  $("sender").textContent = message.sender;
  ($("filename") as HTMLInputElement).value = makeDefaultFilename(message);
}

function renderSelectedFolder() {
  $("selectedFolderPath").textContent = selectedFolder.path === "/" ? "OneDrive" : `OneDrive${selectedFolder.path}`;
}

function renderRecents() {
  const host = $("recentFolders");
  host.innerHTML = "";
  const recents = getRecentFolders();
  $("recentSection").hidden = recents.length === 0;

  for (const folder of recents) {
    const button = document.createElement("button");
    button.className = "recent-folder";
    button.type = "button";
    button.textContent = folder.path === "/" ? "OneDrive" : `OneDrive${folder.path}`;
    button.onclick = () => {
      selectedFolder = folder;
      renderSelectedFolder();
    };
    host.appendChild(button);
  }
}

async function openFolderPicker() {
  browseFolder = selectedFolder || rootFolder;
  browseStack = browseFolder.id === rootFolder.id ? [] : [rootFolder];
  $("picker").hidden = false;
  $("mainView").hidden = true;
  await renderFolderBrowser();
}

function closeFolderPicker() {
  $("picker").hidden = true;
  $("mainView").hidden = false;
}

async function renderFolderBrowser() {
  $("pickerPath").textContent = browseFolder.path === "/" ? "OneDrive" : `OneDrive${browseFolder.path}`;
  $("upButton").toggleAttribute("disabled", browseFolder.id === rootFolder.id);
  const list = $("folderList");
  list.innerHTML = '<div class="loading">Ordner werden geladen …</div>';

  try {
    const folders = await listFolders(browseFolder);
    list.innerHTML = "";
    if (!folders.length) {
      list.innerHTML = '<div class="empty">Dieser Ordner enthält keine Unterordner.</div>';
      return;
    }

    for (const folder of folders) {
      const row = document.createElement("button");
      row.type = "button";
      row.className = "folder-row";
      row.innerHTML = `<span class="folder-icon">▱</span><span>${escapeHtml(folder.name)}</span><span class="chevron">›</span>`;
      row.onclick = async () => {
        browseStack.push(browseFolder);
        browseFolder = folder;
        await renderFolderBrowser();
      };
      list.appendChild(row);
    }
  } catch (error) {
    list.innerHTML = `<div class="error-box">${escapeHtml(errorMessage(error))}</div>`;
  }
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char] || char));
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function saveMessage() {
  setBusy(true);
  setStatus("Nachricht wird als EML abgerufen …");

  try {
    const filenameInput = $("filename") as HTMLInputElement;
    const filename = sanitizeFilename(filenameInput.value);
    filenameInput.value = filename;

    const mime = await getMessageMime(message.graphId);
    setStatus("EML wird nach OneDrive hochgeladen …");
    const saved = await uploadEml(selectedFolder, filename, mime);
    await rememberFolder(selectedFolder).catch(() => undefined);
    renderRecents();

    const savedPath = selectedFolder.path === "/" ? `OneDrive/${saved.name}` : `OneDrive${selectedFolder.path}/${saved.name}`;
    setStatus(`Gespeichert: ${savedPath}`, "success");
  } catch (error) {
    setStatus(errorMessage(error), "error");
  } finally {
    setBusy(false);
  }
}

async function createNewFolder() {
  const name = window.prompt("Name des neuen OneDrive-Ordners:");
  if (!name?.trim()) return;

  try {
    const created = await createFolder(browseFolder, name.trim());
    browseStack.push(browseFolder);
    browseFolder = created;
    await renderFolderBrowser();
  } catch (error) {
    window.alert(errorMessage(error));
  }
}

async function initialize() {
  setBusy(true);
  try {
    message = getCurrentMessage();
    renderMessage();
    await initializeAuth();
    rootFolder = await getOneDriveRoot();
    selectedFolder = getRecentFolders()[0] || rootFolder;
    renderSelectedFolder();
    renderRecents();

    $("chooseFolderButton").onclick = openFolderPicker;
    $("saveButton").onclick = saveMessage;
    $("cancelPickerButton").onclick = closeFolderPicker;
    $("selectFolderButton").onclick = () => {
      selectedFolder = browseFolder;
      renderSelectedFolder();
      closeFolderPicker();
    };
    $("newFolderButton").onclick = createNewFolder;
    $("upButton").onclick = async () => {
      const previous = browseStack.pop();
      if (previous) {
        browseFolder = previous;
        await renderFolderBrowser();
      } else if (browseFolder.id !== rootFolder.id) {
        browseFolder = rootFolder;
        await renderFolderBrowser();
      }
    };
  } catch (error) {
    setStatus(errorMessage(error), "error");
  } finally {
    setBusy(false);
  }
}

Office.onReady(() => initialize());
