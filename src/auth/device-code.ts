import { PublicClientApplication, type AccountInfo, type AuthenticationResult } from "@azure/msal-node";
import { microsoftIdentityNetwork } from "./network";
import { fileCachePlugin } from "./token-cache";
import { AppError } from "../graph/errors";
import { hasExpectedScopes, scopesFor, PERSONAL_AUTHORITY, PERSONAL_TENANT, type AppConfig } from "../config";

export function createPca(config: AppConfig): PublicClientApplication {
  if (!config.clientId) throw new AppError("config_error", "Set MS_TODO_CLIENT_ID to your personal app registration.");
  return new PublicClientApplication({
    auth: { clientId: config.clientId, authority: PERSONAL_AUTHORITY },
    cache: { cachePlugin: fileCachePlugin(config.tokenCachePath) },
    system: { networkClient: microsoftIdentityNetwork(scopesFor(config)), loggerOptions: { piiLoggingEnabled: false, loggerCallback: () => {} } },
  });
}

/** MSAL's scope list comes from the token response. A public client can have other prior grants. */
export function validateGrant(result: AuthenticationResult | null, scopes: string[]): asserts result is AuthenticationResult {
  if (!result?.accessToken || !result.account || result.account.tenantId !== PERSONAL_TENANT ||
      !hasExpectedScopes(result.scopes, scopes)) {
    throw new AppError("permission_denied", "Unexpected account or granted scopes. Revoke consent and review your app registration.");
  }
}

const verificationPages = new Set([
  "https://microsoft.com/link",
  "https://www.microsoft.com/link",
  "https://microsoft.com/devicelogin",
  "https://www.microsoft.com/devicelogin",
]);

export async function acquireByDeviceCode(
  pca: PublicClientApplication,
  scopes: string[],
  onUserCode: (code: string, verificationUri: string) => void,
): Promise<AuthenticationResult> {
  let result;
  try {
    result = await pca.acquireTokenByDeviceCode({ scopes, timeout: 900, deviceCodeCallback: ({ userCode, verificationUri }) => {
      if (typeof userCode !== "string" || !/^[A-Z0-9-]{4,32}$/i.test(userCode)) throw new Error("Invalid code.");
      if (!verificationPages.has(verificationUri)) throw new Error("Unrecognized verification page.");
      onUserCode(userCode, verificationUri);
    } });
  } catch {
    throw new AppError("auth_required", "Microsoft sign-in did not complete. Check the app registration and try again.");
  }
  validateGrant(result, scopes);
  return result;
}

export async function getCachedAccount(pca: PublicClientApplication): Promise<AccountInfo | null> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (!accounts.length) return null;
  const account = accounts[0];
  if (accounts.length !== 1 || account?.tenantId !== PERSONAL_TENANT) {
    throw new AppError("auth_required", "Unexpected cached accounts. Stop the server, log out locally, and sign in again.");
  }
  return account;
}
