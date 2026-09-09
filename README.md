# Outlook EML → OneDrive

Ein kleines Outlook-Add-in für **Outlook auf iOS/Android, Mac, Windows und im Web**. Es speichert die aktuell geöffnete Nachricht als echte `.eml`-Datei in einem vom Benutzer auswählbaren OneDrive-Ordner.

## V1-Funktionen

- Geöffnete E-Mail als vollständige MIME-/EML-Datei abrufen (inkl. Header und Anlagen).
- OneDrive-Ordner im Add-in durchsuchen.
- Zielordner auswählen.
- Neue OneDrive-Ordner direkt aus dem Picker anlegen.
- Die letzten fünf Zielordner über Outlook `RoamingSettings` merken.
- Automatischer, editierbarer Dateiname: `YYYY-MM-DD_HHMM_Absender_Betreff.eml`.
- Responsive Oberfläche für Outlook Mobile und Desktop.
- Kein eigenes Backend: Office.js + MSAL Nested App Authentication + Microsoft Graph.

## Architektur

```text
Outlook (Read Mode)
   │
   ├─ Office.js → ID der aktuell geöffneten Nachricht
   │
   └─ MSAL NAA → Microsoft Graph Token
                  │
                  ├─ GET /me/messages/{id}/$value
                  │      → MIME/EML
                  │
                  ├─ GET /me/drive/items/{folder}/children
                  │      → Folder Picker
                  │
                  └─ PUT /me/drive/items/{folder}:/{name}:/content
                         → OneDrive
```

## 1. Entra-App registrieren

1. Microsoft Entra Admin Center / Azure Portal → **App registrations** → **New registration**.
2. Für einen rein internen Tenant empfiehlt sich **Accounts in this organizational directory only**.
3. Unter **Authentication** eine **Single-page application (SPA)** konfigurieren.
4. Für lokale Entwicklung diese Redirect URIs hinzufügen:
   - `brk-multihub://localhost:3000`
   - `https://localhost:3000/auth.html`
5. Für Produktion zusätzlich den NAA-Broker-Redirect für deinen Hosting-Origin hinzufügen, z. B.:
   - `brk-multihub://outlook-addins.example.org`
   - `https://outlook-addins.example.org/auth.html`
6. Unter **API permissions → Microsoft Graph → Delegated permissions** hinzufügen:
   - `Mail.Read`
   - `Files.ReadWrite`
7. Optional: Admin consent tenantweit erteilen, damit Benutzer nicht einzeln zustimmen müssen.
8. Die **Application (client) ID** kopieren.

## 2. Client-ID eintragen

In `src/taskpane/config.ts`:

```ts
export const CLIENT_ID = "DEINE-CLIENT-ID";
```

Es wird **kein Client Secret** benötigt und keines darf in das Add-in eingebaut werden.

## 3. Lokal starten

Voraussetzungen: Node.js und npm.

```bash
npm install
npm run start
```

`office-addin-debugging` erzeugt für localhost ein Entwicklungszertifikat und startet den HTTPS-Dev-Server.

Falls du nur den Webserver starten willst:

```bash
npm run dev-server
```

## 4. iPhone testen

Das Manifest enthält `MobileMessageReadCommandSurface`. Für einen Tenant-Test ist der sauberste Weg, die Custom App/Manifest-Datei zentral einem Testbenutzer bereitzustellen. Sobald Outlook Mobile die Bereitstellung synchronisiert hat, erscheint das Add-in in einer **geöffneten Nachricht** im Add-in-/Apps-Menü.

Wichtig: Outlook Mobile unterstützt für normale Add-ins primär **Message Read Mode**. Genau deshalb wird die Mail erst geöffnet und dann über **Als EML speichern** archiviert.

## 5. Produktion bauen

Das Add-in braucht statisches HTTPS-Hosting. Beispiel mit einem beliebigen Webhost/Azure Static Web Apps:

```bash
ADDIN_PUBLIC_URL=https://outlook-addins.example.org npm run build
```

Der Build liegt anschließend in `dist/`. `dist/manifest.xml` enthält automatisch die Produktions-URLs.

Danach:

1. `dist/` auf dem HTTPS-Host veröffentlichen.
2. Prüfen, dass die Entra-App den passenden Broker-Redirect enthält:
   `brk-multihub://outlook-addins.example.org`
3. `dist/manifest.xml` als Custom Outlook Add-in verteilen.

## Berechtigungen

### Graph

- `Mail.Read`: erforderlich, um die MIME-Repräsentation der geöffneten Mail via `/$value` zu lesen.
- `Files.ReadWrite`: erforderlich, um OneDrive-Ordner aufzulisten, Ordner anzulegen und die `.eml`-Datei hochzuladen.

### Outlook-Manifest

Das Manifest verwendet nur `ReadItem`. Der Zugriff auf andere Mailbox-/OneDrive-Daten erfolgt separat und benutzerdelegiert über Microsoft Graph.

## Mobile-ID-Behandlung

Outlook Mobile liefert `item.itemId` bereits im REST-/Graph-kompatiblen Format. Auf Mac/Windows/Web wird die Outlook-ID via `convertToRestId(..., v2_0)` konvertiert. Das ist in `src/taskpane/outlook.ts` gekapselt.

## Dateigröße

Der einfache Graph-Upload unterstützt Dateien bis 250 MB. Exchange-/Outlook-Nachrichten liegen typischerweise darunter. Falls das Add-in später auch sehr große EML-Dateien archivieren soll, kann `uploadEml()` auf Graph Upload Sessions erweitert werden.

## Sinnvolle V2

- Favoritenordner zusätzlich zu „Zuletzt verwendet“.
- SharePoint-Dokumentbibliotheken neben OneDrive.
- Namensschema konfigurierbar machen.
- Kollisionen explizit mit „Umbenennen / Ersetzen“ behandeln.
- Mehrere markierte Nachrichten in einem Durchgang archivieren (wo vom Client unterstützt).
- Optional eine kleine Metadaten-Datei oder SharePoint-Spalten ergänzen.
