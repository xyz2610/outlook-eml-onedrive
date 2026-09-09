# Outlook EML Archive

Save the open Outlook message as an EML file in a selected OneDrive folder.

Hosting: https://xyz2610.github.io/outlook-eml-onedrive/

GitHub Pages source must be set to **GitHub Actions**. The workflow installs dependencies, checks TypeScript, builds the static files, validates the production manifest, and deploys `dist`.

Microsoft Entra: single-tenant SPA registration with delegated `Mail.Read` and `Files.ReadWrite`. NAA redirect: `brk-multihub://xyz2610.github.io`. For browser fallback, also register `https://xyz2610.github.io/outlook-eml-onedrive/auth.html` as an SPA redirect.

Application and tenant IDs in the source are public identifiers. This application uses no client secret. Never commit tokens, passwords, or mailbox content.

Once deployed, download manifest.xml from the hosted site and install through https://aka.ms/olksideload using My add-ins > Add a custom add-in > Add from File. Test with your own Microsoft 365 mailbox first, then the same account in Outlook iOS.

The current deployment package has had local structural checks only. Dependency installation was blocked locally by network restrictions; the GitHub workflow must pass before installation. End-to-end Outlook and OneDrive behavior remains to be tested.
