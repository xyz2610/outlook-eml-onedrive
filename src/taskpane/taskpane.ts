import { initializeLanguage, t } from "./i18n";
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
let busy = false;
let ready = false;
let savedSuccessfully = false;
let folderRequest = 0;
let creatingFolder = false;

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

function setStatus(text: string, kind: "info" | "success" | "error" = "info") {
  const el = $("status");
  el.textContent = text;
  el.dataset.kind = kind;
  el.hidden = !text;
}

function setBusy(value: boolean) {
  busy = value;
  for (const id of ["saveButton", "chooseFolderButton", "selectFolderButton", "newFolderButton", "upButton", "cancelPickerButton"]) {
    $(id).toggleAttribute("disabled", busy || !ready || savedSuccessfully);
  }
  if (rootFolder && browseFolder) $("upButton").toggleAttribute("disabled", busy || browseFolder.id === rootFolder.id);
  document.body.dataset.busy = busy ? "true" : "false";
}

function closeAddin() {
  try {
    if (typeof Office.context.ui.closeContainer === "function") {
      Office.context.ui.closeContainer();
    }
  } catch { /* Preserve the successful save even if this host cannot close. */ }
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
      if (busy || savedSuccessfully) return;
      selectedFolder = folder;
      renderSelectedFolder();
    };
    host.appendChild(button);
  }
}

async function openFolderPicker() {
  if (busy || !ready) return;
  browseFolder = selectedFolder || rootFolder;
  browseStack = browseFolder.id === rootFolder.id ? [] : [rootFolder];
  $("picker").hidden = false;
  $("mainView").hidden = true;
  $("newFolderForm").hidden = true;
  window.scrollTo(0, 0);
  await renderFolderBrowser();
}

function closeFolderPicker() {
  $("picker").hidden = true;
  $("mainView").hidden = false;
  $("mainContent").scrollTop = 0;
  window.scrollTo(0, 0);
}

async function renderFolderBrowser() {
  const request = ++folderRequest;
  const parent = browseFolder;
  $("pickerPath").textContent = browseFolder.path === "/" ? "OneDrive" : `OneDrive${browseFolder.path}`;
  $("upButton").toggleAttribute("disabled", browseFolder.id === rootFolder.id);
  const list = $("folderList");
  list.scrollTop = 0;
  list.innerHTML = `<div class="loading">${escapeHtml(t("loadingFolders"))}</div>`;

  try {
    const folders = await listFolders(parent);
    if (request !== folderRequest) return;
    list.innerHTML = "";
    if (!folders.length) {
      list.innerHTML = `<div class="empty">${escapeHtml(t("emptyFolder"))}</div>`;
      return;
    }

    for (const folder of folders) {
      const row = document.createElement("button");
      row.type = "button";
      row.className = "folder-row";
      row.innerHTML = `<span class="folder-icon">▱</span><span>${escapeHtml(folder.name)}</span><span class="chevron">›</span>`;
      row.onclick = async () => {
        if (busy || creatingFolder) return;
        browseStack.push(browseFolder);
        browseFolder = folder;
        await renderFolderBrowser();
      };
      list.appendChild(row);
    }
  } catch (error) {
    if (request !== folderRequest) return;
    list.innerHTML = `<div class="error-box">${escapeHtml(errorMessage(error))}</div>`;
  }
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char] || char));
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : t("requestFailed");
}

async function saveMessage() {
  if (busy || !ready || savedSuccessfully) return;
  setBusy(true);
  setStatus(t("readingMessage"));

  try {
    const filenameInput = $("filename") as HTMLInputElement;
    const filename = sanitizeFilename(filenameInput.value);
    filenameInput.value = filename;

    const mime = await getMessageMime(message.graphId);
    setStatus(t("uploading"));
    const saved = await uploadEml(selectedFolder, filename, mime);
    savedSuccessfully = true;
    await Promise.race([rememberFolder(selectedFolder).catch(() => undefined), new Promise<void>(resolve => setTimeout(resolve, 1200))]);
    renderRecents();

    const savedPath = selectedFolder.path === "/" ? `OneDrive/${saved.name}` : `OneDrive${selectedFolder.path}/${saved.name}`;
    setStatus(t("saved", { path: savedPath }), "success");
    $("saveButton").hidden = true;
    $("closeButton").hidden = false;
    setTimeout(closeAddin, 650);
  } catch (error) {
    setStatus(errorMessage(error), "error");
  } finally {
    setBusy(false);
  }
}

function showNewFolderForm() {
  if (busy || creatingFolder) return;
  $("newFolderForm").hidden = false;
  $("folderError").hidden = true;
  ($("newFolderName") as HTMLInputElement).value = "";
  ($("newFolderName") as HTMLInputElement).focus();
}

async function createNewFolder(event: Event) {
  event.preventDefault();
  if (creatingFolder || busy) return;
  const input = $("newFolderName") as HTMLInputElement;
  const name = input.value.trim();
  if (!name) { input.focus(); return; }
  creatingFolder = true;
  const parent = browseFolder;
  $("folderError").hidden = true;
  for (const id of ["createFolderButton", "cancelNewFolderButton", "cancelPickerButton", "selectFolderButton", "upButton", "newFolderButton", "newFolderName"]) $(id).toggleAttribute("disabled", true);
  $("createFolderButton").textContent = t("creating");
  try {
    const created = await createFolder(parent, name);
    browseStack.push(parent);
    browseFolder = created;
    $("newFolderForm").hidden = true;
    input.blur();
    await renderFolderBrowser();
    $("selectFolderButton").focus();
  } catch (error) {
    $("folderError").textContent = errorMessage(error);
    $("folderError").hidden = false;
  } finally {
    creatingFolder = false;
    for (const id of ["createFolderButton", "cancelNewFolderButton", "newFolderName"]) $(id).removeAttribute("disabled");
    $("createFolderButton").textContent = t("create");
    setBusy(false);
  }
}

async function initialize() {
  setBusy(true);
  try {
    message = getCurrentMessage();
    renderMessage();
    await initializeAuth();
    rootFolder = await getOneDriveRoot();
    ready = true;
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
      void saveMessage();
    };
    $("newFolderButton").onclick = showNewFolderForm;
    $("newFolderForm").onsubmit = createNewFolder;
    $("cancelNewFolderButton").onclick = () => { $("newFolderForm").hidden = true; $("newFolderButton").focus(); };
    $("closeButton").onclick = closeAddin;
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

Office.onReady(() => {
  initializeLanguage(Office.context.displayLanguage, navigator.language);
  return initialize();
});
