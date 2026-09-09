import { t } from "./i18n";
import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
  type IPublicClientApplication,
  type AuthenticationResult,
} from "@azure/msal-browser";
import { CLIENT_ID, TENANT_ID } from "./config";

let client: IPublicClientApplication | undefined;

function getMsalConfig() {
  if (!CLIENT_ID || CLIENT_ID.startsWith("YOUR-")) {
    throw new Error(t("missingClient"));
  }

  return {
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      redirectUri: new URL("auth.html", window.location.href).href,
      postLogoutRedirectUri: new URL("auth.html", window.location.href).href,
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
  if (!client) throw new Error(t("authInitFailed"));

  const active = client.getActiveAccount() || client.getAllAccounts()[0];
  if (active && !client.getActiveAccount()) client.setActiveAccount(active);

  try {
    const silent = await client.acquireTokenSilent({ scopes, ...(active ? { account: active } : {}) });
    return silent.accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw new Error(t("authFailed"));
  }

  let result: AuthenticationResult;
  try {
    result = await client.acquireTokenPopup({ scopes, ...(active ? { account: active } : {}) });
  } catch (error) {
    throw new Error(
      t("authFailed")
    );
  }

  if (result.account) client.setActiveAccount(result.account);
  return result.accessToken;
}
