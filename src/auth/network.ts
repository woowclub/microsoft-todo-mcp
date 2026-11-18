import type { INetworkModule, NetworkRequestOptions, NetworkResponse } from "@azure/msal-node";
import { hasExpectedScopes, PERSONAL_AUTHORITY, PERSONAL_TENANT } from "../config";
import { AppError } from "../graph/errors";
import { readBoundedText } from "../lib/http-body";

const origin = new URL(PERSONAL_AUTHORITY).origin;
const getPaths = new Set(["/common/discovery/instance"]);
const postPaths = new Set<string>();
for (const tenant of ["consumers", PERSONAL_TENANT]) {
  getPaths.add(`/${tenant}/v2.0/.well-known/openid-configuration`);
  postPaths.add(`/${tenant}/oauth2/v2.0/token`);
  postPaths.add(`/${tenant}/oauth2/v2.0/devicecode`);
}

const sdkMetrics = /^x-client-(?:sku|ver|os|cpu|current-telemetry|last-telemetry)$/i;

// MSAL performs discovery and sends refresh tokens and device codes in POST bodies.
// A redirect could replay those bodies. Apply the same boundary after discovery.
export function microsoftIdentityNetwork(scopes: string[], fetchImpl: typeof fetch = fetch): INetworkModule {
  async function request<T>(method: "GET" | "POST", raw: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
    const failed = new AppError("auth_required", "Microsoft identity request failed.");
    try {
      const url = new URL(raw);
      const paths = method === "GET" ? getPaths : postPaths;
      if (url.origin !== origin || url.username || url.password || url.hash || !paths.has(url.pathname)) throw failed;
      const body = new URLSearchParams(options?.body);
      const submittedScopes = body.get("scope");
      if (method === "POST" && (!submittedScopes || !hasExpectedScopes(submittedScopes.split(/\s+/), scopes))) throw failed;
      const headers = new Headers(options?.headers);
      for (const key of [...url.searchParams.keys(), ...body.keys(), ...headers.keys()]) {
        if (sdkMetrics.test(key)) {
          url.searchParams.delete(key);
          body.delete(key);
          headers.delete(key);
        }
      }
      const response = await fetchImpl(url, {
        method, headers, body: method === "POST" ? body.toString() : undefined,
        redirect: "error", signal: AbortSignal.timeout(15_000),
      });
      if (response.status >= 300 && response.status < 400) {
        await response.body?.cancel();
        throw failed;
      }
      const data = JSON.parse(await readBoundedText(response, 1024 * 1024, failed));
      if (!data || typeof data !== "object" || Array.isArray(data)) throw failed;
      // Validate the grant before MSAL can persist an unexpectedly privileged token.
      // OAuth permits omission when the granted scope matches the request.
      if (response.ok && url.pathname.endsWith("/devicecode") &&
          (!Number.isInteger(data.interval) || data.interval < 1 || data.interval > 60 ||
           !Number.isInteger(data.expires_in) || data.expires_in < 1 || data.expires_in > 3600)) throw failed;
      if (response.ok && url.pathname.endsWith("/token")) {
        if (typeof data.access_token !== "string" || !data.access_token) throw failed;
        if (data.scope !== undefined &&
            (typeof data.scope !== "string" || !hasExpectedScopes(data.scope.split(/\s+/), scopes))) throw failed;
      }
      return {
        status: response.status,
        headers: Object.fromEntries(response.headers.entries()),
        body: data as T,
      };
    } catch { throw failed; }
  }
  return {
    sendGetRequestAsync: (url, options) => request("GET", url, options),
    sendPostRequestAsync: (url, options) => request("POST", url, options),
  };
}
