# Als EML speichern / Save as EML — 1.1.0

Outlook add-in to save the current email as an EML file in a selected OneDrive folder.

The user interface follows Office.context.displayLanguage. German language variants use German; English and other languages use English. Browser language is used only when Outlook does not provide a display language. Email content, sender names, and existing folder names are preserved.

Manifest version: 1.1.0.0. English defaults with German regional overrides for the app name, ribbon labels, and descriptions. Same add-in ID and delegated Graph permissions as before (Mail.Read, Files.ReadWrite).

## Deploy this update

Upload the contents of the update folder to the root of the existing GitHub repository. Keep the existing .github/workflows/deploy.yml. GitHub Actions runs typecheck, build, production-manifest validation, and Pages deployment.

After the workflow succeeds, update the existing EML Archive app in Microsoft 365 admin center > Settings > Integrated apps using the new production manifest. The manifest must be updated for Outlook to pick up the localized name. Do not create a second add-in. Keep the existing user assignments.

Hosting: https://xyz2610.github.io/outlook-eml-onedrive/

Local checks: translation parity and keys, DE/AT/CH/US/GB and fallback language selection, static/accessibility labels, folder creation, save/close behavior, Graph error messages, manifest XML and unchanged app identity/permissions. Outlook and OneDrive were simulated; an end-to-end test of this localized release in Outlook remains necessary. Local dependency downloads are restricted, so full build/schema validation takes place in GitHub Actions.
