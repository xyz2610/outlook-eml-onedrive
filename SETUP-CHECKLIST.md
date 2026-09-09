# Setup-Checkliste

1. Entra App Registration anlegen.
2. SPA Redirects setzen:
   - `brk-multihub://localhost:3000`
   - `https://localhost:3000/auth.html`
3. Delegierte Graph-Berechtigungen hinzufügen:
   - `Mail.Read`
   - `Files.ReadWrite`
4. Client-ID in `src/taskpane/config.ts` einsetzen.
5. `npm install`
6. `npm run start`
7. In Outlook Desktop/Web mit einer geöffneten Mail testen.
8. Für iOS das Manifest als Custom App an einen Testbenutzer verteilen.
9. Für Produktion statisches HTTPS-Hosting bereitstellen.
10. Produktions-Broker-Redirect ergänzen, z. B. `brk-multihub://outlook-addins.example.org`.
11. `ADDIN_PUBLIC_URL=https://outlook-addins.example.org npm run build`
12. `dist/manifest.xml` zentral verteilen.
