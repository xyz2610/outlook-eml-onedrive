import {
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  type AuthenticationResult,
} from "@azure/msal-browser";
import { CLIENT_ID } from "./config";

let client: IPublicClientApplication | undefined;

function getMsalConfig() {
  if (!CLIENT_ID || CLIENT_ID.startsWith("YOUR-")) {
    throw new Error("Bitte zuerst CLIENT_ID in src/taskpane/config.ts setzen.");
  }

  return {
    auth: {
      clientId: CLIENT_ID,
      redirectUri: `${window.location.origin}/auth.html`,
      postLogoutRedirectUri: `${window.location.origin}/auth.html`,
      clientCapabilities: ["CP1"],
    },
    cache: {
      cacheLocation: "localStorage" as const,
    },
  };
}

export async function initializeAuth(): Promise<void> {
  if (client) return;
  client = await createNestablePublicClientApplication(getMsalConfig());
}

export async function getAccessToken(scopes: string[]): Promise<string> {
  await initializeAuth();
  if (!client) throw new Error("Authentifizierung konnte nicht initialisiert werden.");

  const active = client.getActiveAccount() || client.getAllAccounts()[0];
  if (active && !client.getActiveAccount()) client.setActiveAccount(active);

  if (active) {
    try {
      const silent = await client.acquireTokenSilent({ scopes, account: active });
      return silent.accessToken;
    } catch {
      // Continue with the brokered interactive flow below.
    }
  }

  let result: AuthenticationResult;
  try {
    result = await client.acquireTokenPopup({ scopes, prompt: active ? undefined : "select_account" });
  } catch (error) {
    throw new Error(
      `Microsoft-Anmeldung fehlgeschlagen. Prüfe NAA/Redirect-URI und die Graph-Berechtigungen. ${String(error)}`
    );
  }

  if (result.account) client.setActiveAccount(result.account);
  return result.accessToken;
}
