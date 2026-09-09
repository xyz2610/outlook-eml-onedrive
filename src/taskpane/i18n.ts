/** Outlook display language takes precedence over the browser language. */
export const en = {
  "appName": "Save as EML",
  "filename": "File name",
  "location": "Save to",
  "recent": "Recent folders",
  "saveClose": "Save & close",
  "close": "Close",
  "cancel": "Cancel",
  "chooseFolder": "Choose a folder",
  "up": "Go up one level",
  "newFolder": "New folder",
  "folderName": "Folder name",
  "create": "Create",
  "saveHere": "Save here",
  "loadingFolders": "Loading folders …",
  "emptyFolder": "This folder has no subfolders.",
  "readingMessage": "Retrieving email as EML …",
  "uploading": "Saving EML to OneDrive …",
  "saved": "Saved: {path}",
  "creating": "Creating …",
  "unknownSender": "Unknown sender",
  "noSubject": "No subject",
  "fallbackFilename": "Message",
  "missingMessage": "No ID is available for the open email.",
  "missingClient": "The Microsoft sign-in configuration is incomplete.",
  "authInitFailed": "Microsoft sign-in could not be initialized.",
  "authFailed": "Microsoft sign-in failed. Please try again.",
  "settingsFailed": "Your recent folders could not be saved.",
  "requestFailed": "The request failed. Please try again.",
  "networkFailed": "Unable to connect. Check your connection and try again.",
  "accessDenied": "Access denied. Check that you have permission to access this email or folder.",
  "notFound": "The email or folder could not be found.",
  "throttled": "Microsoft is receiving too many requests. Please try again shortly.",
  "serviceUnavailable": "Microsoft is temporarily unavailable. Please try again shortly.",
  "sessionExpired": "Your sign-in session has expired. Close the add-in and open it again.",
  "invalidRequest": "The request could not be completed. Check the file or folder name.",
  "nameConflict": "An item with this name already exists. Choose a different name."
} as const;
export type TranslationKey = keyof typeof en;
export const de: Record<TranslationKey, string> = {
  "appName": "Als EML speichern",
  "filename": "Dateiname",
  "location": "Speicherort",
  "recent": "Zuletzt verwendet",
  "saveClose": "Speichern & schließen",
  "close": "Schließen",
  "cancel": "Abbrechen",
  "chooseFolder": "Ordner wählen",
  "up": "Eine Ebene nach oben",
  "newFolder": "Neuer Ordner",
  "folderName": "Ordnername",
  "create": "Erstellen",
  "saveHere": "Hier speichern",
  "loadingFolders": "Ordner werden geladen …",
  "emptyFolder": "Dieser Ordner enthält keine Unterordner.",
  "readingMessage": "Nachricht wird als EML abgerufen …",
  "uploading": "EML wird in OneDrive gespeichert …",
  "saved": "Gespeichert: {path}",
  "creating": "Wird erstellt …",
  "unknownSender": "Unbekannter Absender",
  "noSubject": "Ohne Betreff",
  "fallbackFilename": "Nachricht",
  "missingMessage": "Für die geöffnete Nachricht ist keine ID verfügbar.",
  "missingClient": "Die Konfiguration der Microsoft-Anmeldung ist unvollständig.",
  "authInitFailed": "Die Microsoft-Anmeldung konnte nicht initialisiert werden.",
  "authFailed": "Die Microsoft-Anmeldung ist fehlgeschlagen. Bitte versuche es erneut.",
  "settingsFailed": "Die zuletzt verwendeten Ordner konnten nicht gespeichert werden.",
  "requestFailed": "Die Anfrage ist fehlgeschlagen. Bitte versuche es erneut.",
  "networkFailed": "Keine Verbindung möglich. Prüfe deine Verbindung und versuche es erneut.",
  "accessDenied": "Zugriff verweigert. Prüfe deine Berechtigung für diese Nachricht oder diesen Ordner.",
  "notFound": "Die Nachricht oder der Ordner wurde nicht gefunden.",
  "throttled": "Microsoft erhält zu viele Anfragen. Bitte versuche es in Kürze erneut.",
  "serviceUnavailable": "Microsoft ist vorübergehend nicht erreichbar. Bitte versuche es in Kürze erneut.",
  "sessionExpired": "Deine Anmeldung ist abgelaufen. Schließe das Add-in und öffne es erneut.",
  "invalidRequest": "Die Anfrage konnte nicht ausgeführt werden. Prüfe den Datei- oder Ordnernamen.",
  "nameConflict": "Ein Eintrag mit diesem Namen ist bereits vorhanden. Wähle einen anderen Namen."
};
export type Language = "de" | "en";
let language: Language = "en";

export function selectLanguage(displayLanguage?: string, browserLanguage?: string): Language {
  const locale = (displayLanguage || browserLanguage || "en").trim().replace(/_/g, "-").toLowerCase();
  return locale === "de" || locale.startsWith("de-") ? "de" : "en";
}

export function initializeLanguage(displayLanguage?: string, browserLanguage?: string): void {
  language = selectLanguage(displayLanguage, browserLanguage);
  document.documentElement.lang = language;
  document.title = t("appName");
  document.querySelectorAll<HTMLElement>("[data-i18n]").forEach(element => {
    element.textContent = t(element.dataset.i18n as TranslationKey);
  });
  document.querySelectorAll<HTMLInputElement>("[data-i18n-placeholder]").forEach(element => {
    element.placeholder = t(element.dataset.i18nPlaceholder as TranslationKey);
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-label]").forEach(element => {
    element.setAttribute("aria-label", t(element.dataset.i18nLabel as TranslationKey));
  });
}

export function t(key: TranslationKey, values: Record<string, string> = {}): string {
  const text = (language === "de" ? de : en)[key] || en[key];
  return text.replace(/\{(\w+)\}/g, (placeholder, name: string) => values[name] ?? placeholder);
}
