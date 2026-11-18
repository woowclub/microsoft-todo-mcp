import { mkdtemp, realpath, rm } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import type { PublicClientApplication } from "@azure/msal-node";
import { AuthManager } from "../src/auth/auth-state";
import { acquireByDeviceCode, createPca, validateGrant } from "../src/auth/device-code";
import { loadConfig, PERSONAL_AUTHORITY, PERSONAL_TENANT, scopesFor } from "../src/config";
import { createLogger } from "../src/lib/logger";

const scopes = ["Tasks.ReadWrite"];
const account = { username: "private@example.test", homeAccountId: "home-1", tenantId: PERSONAL_TENANT };
const good = () => ({ account, accessToken: "test-only", expiresOn: new Date(Date.now() + 3_600_000), scopes });
function fixture(accounts: object[] = [], silent = async (): Promise<unknown> => good()) {
  const acquireTokenSilent = vi.fn(silent);
  const acquireTokenByDeviceCode = vi.fn();
  const pca = {
    getTokenCache: () => ({ getAllAccounts: async () => accounts }),
    acquireTokenSilent,
    acquireTokenByDeviceCode,
  } as unknown as PublicClientApplication;
  const auth = new AuthManager(pca, scopes, createLogger("silent"));
  return { auth, pca, acquireTokenSilent, acquireTokenByDeviceCode };
}

describe("personal-account authentication", () => {
  it("requires a dedicated client ID and fixes the authority and task scopes", () => {
    const config = loadConfig({});
    expect(config.clientId).toBe("");
    expect(config.authority).toBe(PERSONAL_AUTHORITY);
    expect(scopesFor(config)).toEqual(["Tasks.ReadWrite"]);
    const ro = loadConfig({ MS_TODO_SCOPE_READONLY: "1" });
    expect(scopesFor(ro)).toEqual(["Tasks.Read"]);
    expect(ro.tokenCachePath).not.toBe(config.tokenCachePath);
    expect(() => loadConfig({ MS_TODO_CLIENT_ID: "../client" })).toThrow();
    expect(() => loadConfig({ MS_TODO_AUTHORITY: "https://example.test" })).toThrow();
    expect(() => loadConfig({ MS_TODO_TOKEN_CACHE: "/tmp/shared.json" })).toThrow();
  });

  it.each([
    "https://microsoft.com/link", "https://www.microsoft.com/link",
    "https://microsoft.com/devicelogin", "https://www.microsoft.com/devicelogin",
  ])("preserves the issued verification page %s", async (verificationUri) => {
    const { pca, acquireTokenByDeviceCode } = fixture();
    acquireTokenByDeviceCode.mockImplementation(async ({ deviceCodeCallback }) => {
      deviceCodeCallback({ userCode: "TESTONLY", verificationUri });
      return good();
    });
    const prompt = vi.fn();
    await acquireByDeviceCode(pca, scopes, prompt);
    expect(prompt).toHaveBeenCalledWith("TESTONLY", verificationUri);
  });

  it.each([undefined, "http://microsoft.com/link", "https://microsoft.com.example/link",
    "https://microsoft.com/link?next=private", "https://microsoft.com/unrecognized",
  ])("suppresses the prompt for an unsupported verification page", async (verificationUri) => {
    const { pca, acquireTokenByDeviceCode } = fixture();
    acquireTokenByDeviceCode.mockImplementation(async ({ deviceCodeCallback }) => {
      deviceCodeCallback({ userCode: "TESTONLY", verificationUri });
      return good();
    });
    const prompt = vi.fn();
    await expect(acquireByDeviceCode(pca, scopes, prompt)).rejects.toMatchObject({ code: "auth_required" });
    expect(prompt).not.toHaveBeenCalled();
  });

  it("never starts an interactive flow for an MCP request", async () => {
    const { auth, acquireTokenByDeviceCode } = fixture();
    await auth.init();
    await expect(auth.getAccessToken()).rejects.toMatchObject({ code: "auth_required" });
    expect(auth.status()).toEqual({ state: "uninitialized" });
    expect(acquireTokenByDeviceCode).not.toHaveBeenCalled();
  });

  it("shares concurrent refreshes and caches only a validated unexpired token", async () => {
    const { auth, acquireTokenSilent } = fixture([account]);
    await auth.init();
    expect(await Promise.all([auth.getAccessToken(), auth.getAccessToken()])).toEqual(["test-only", "test-only"]);
    expect(await auth.getAccessToken()).toBe("test-only");
    expect(acquireTokenSilent).toHaveBeenCalledTimes(1);
    expect(auth.status()).toEqual({ state: "authenticated", tokenExpiresOn: expect.any(String) });
    expect(auth.accountId()).toBe(account.homeAccountId);
  });

  it.each([[account, account], [{ ...account, tenantId: "unrelated" }]])("rejects ambiguous or non-personal accounts", async (...accounts) => {
    const { auth } = fixture(accounts);
    await expect(auth.init()).rejects.toMatchObject({ code: "auth_required" });
  });

  it("rejects extra granted Graph scopes and changed account identity", async () => {
    for (const bad of [{ ...good(), scopes: [...scopes, "Mail.Read"] }, { ...good(), account: { ...account, homeAccountId: "other" } }]) {
      const { auth } = fixture([account], async () => bad);
      await auth.init();
      await expect(auth.getAccessToken()).rejects.toMatchObject({ code: "auth_required" });
      expect(auth.status()).toEqual({ state: "refresh_failed" });
    }
    const { pca, acquireTokenByDeviceCode } = fixture();
    acquireTokenByDeviceCode.mockResolvedValue({ ...good(), scopes: ["Files.Read"] });
    await expect(acquireByDeviceCode(pca, scopes, vi.fn())).rejects.toMatchObject({ code: "permission_denied" });
    expect(() => validateGrant(null, scopes)).toThrow();
  });

  it("keeps provider errors and codes out of failures", async () => {
    const { auth, pca, acquireTokenByDeviceCode } = fixture([account], async () => { throw new Error("test-only-secret"); });
    acquireTokenByDeviceCode.mockRejectedValue(new Error("test-only-secret"));
    await auth.init();
    for (const op of [() => auth.getAccessToken(), () => acquireByDeviceCode(pca, scopes, vi.fn())]) {
      await expect(op()).rejects.not.toThrow(/test-only-secret/);
    }
  });

  it("reports configuration errors without fetching credentials", async () => {
    const auth = new AuthManager(null, scopes, createLogger("silent"), undefined, "Set MS_TODO_CLIENT_ID.");
    await expect(auth.getAccessToken()).rejects.toMatchObject({ code: "config_error" });
  });
});

describe("Microsoft identity transport", () => {
  it("pins discovery and credential POSTs to Microsoft and disables redirects", async () => {
    const { microsoftIdentityNetwork } = await import("../src/auth/network");
    const transport = vi.fn<typeof fetch>().mockResolvedValue(Response.json({ access_token: "fixture" }));
    const network = microsoftIdentityNetwork(scopes, transport);
    const body = "scope=Tasks.ReadWrite&refresh_token=fake&x-client-last-telemetry=fixture";
    for (const url of ["https://example.test/consumers/oauth2/v2.0/token", PERSONAL_AUTHORITY + "/other", PERSONAL_AUTHORITY.replace("https", "http") + "/oauth2/v2.0/token"]) {
      await expect(network.sendPostRequestAsync(url, { body })).rejects.toMatchObject({ code: "auth_required" });
    }
    expect(transport).not.toHaveBeenCalled();
    await network.sendPostRequestAsync(PERSONAL_AUTHORITY + "/oauth2/v2.0/token", { body });
    expect(transport).toHaveBeenCalledWith(new URL(PERSONAL_AUTHORITY + "/oauth2/v2.0/token"), expect.objectContaining({
      body: "scope=Tasks.ReadWrite&refresh_token=fake", method: "POST", redirect: "error", signal: expect.any(AbortSignal),
    }));
    transport.mockResolvedValueOnce(new Response(null, { status: 307, headers: { location: "https://example.test/" } }));
    await expect(network.sendPostRequestAsync(PERSONAL_AUTHORITY + "/oauth2/v2.0/token", { body })).rejects.toThrow("Microsoft identity request failed.");
    expect(transport).toHaveBeenCalledTimes(2);
  });
});

describe("MSAL grant handling", () => {
  it("rejects an expanded grant before the SDK can persist the response", async () => {
    const { microsoftIdentityNetwork } = await import("../src/auth/network");
    const transport = vi.fn<typeof fetch>().mockResolvedValue(Response.json({ access_token: "test-only", scope: "Tasks.ReadWrite Files.Read" }));
    const network = microsoftIdentityNetwork(scopes, transport);
    await expect(network.sendPostRequestAsync(PERSONAL_AUTHORITY + "/oauth2/v2.0/token", {
      body: "scope=Tasks.ReadWrite&refresh_token=fake",
    })).rejects.toMatchObject({ code: "auth_required" });
    await expect(network.sendPostRequestAsync(PERSONAL_AUTHORITY + "/oauth2/v2.0/devicecode", {
      body: "scope=Mail.Read",
    })).rejects.toMatchObject({ code: "auth_required" });
    expect(transport).toHaveBeenCalledTimes(1);
  });
});

it("applies the transport and expected OAuth scopes on an actual MSAL refresh request", async () => {
  const dir = await mkdtemp(join(await realpath("/tmp"), "todo-msal-"));
  const transport = vi.fn<typeof fetch>().mockImplementation(async () =>
    Response.json({ error: "invalid_grant", error_description: "synthetic rejection" }, { status: 400 }));
  vi.stubGlobal("fetch", transport);
  try {
    const config = loadConfig({ MS_TODO_CLIENT_ID: "00000000-0000-4000-8000-000000000001" });
    const pca = createPca({ ...config, tokenCachePath: join(dir, "fixture.json") });
    await expect(pca.acquireTokenByRefreshToken({ refreshToken: "fixture-only", scopes: scopesFor(config) })).rejects.toThrow();
    const post = transport.mock.calls.filter(([, options]) => options?.method === "POST");
    expect(post).toHaveLength(1);
    const [url, options] = post[0]!;
    expect(new URL(String(url)).pathname).toBe("/consumers/oauth2/v2.0/token");
    expect(new URL(String(url)).origin).toBe("https://login.microsoftonline.com");
    expect(options?.redirect).toBe("error");
    const body = new URLSearchParams(String(options?.body));
    expect(body.get("refresh_token")).toBe("fixture-only");
    expect(body.get("scope")?.split(" ").sort()).toEqual(["Tasks.ReadWrite", "offline_access", "openid", "profile"].sort());
    expect([...body.keys()].some((k) => k.includes("telemetry"))).toBe(false);
  } finally {
    vi.unstubAllGlobals();
    await rm(dir, { recursive: true, force: true });
  }
});
