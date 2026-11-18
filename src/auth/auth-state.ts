import type { AccountInfo, PublicClientApplication } from "@azure/msal-node";
import { getCachedAccount, validateGrant } from "./device-code";
import { AppError } from "../graph/errors";
import type { TokenProvider } from "../graph/client";
import type { Logger } from "../lib/logger";

type AuthState = "uninitialized" | "token_expired" | "authenticated" | "refresh_failed";

export const LOGIN_HINT = "Run the built local login command in your own terminal, then restart the MCP server.";

export class AuthManager implements TokenProvider {
  private state: AuthState = "uninitialized";
  private account: AccountInfo | null = null;
  private token: { value: string; expiresOn: Date } | null = null;
  private refreshing: Promise<string> | null = null;

  constructor(
    private readonly pca: PublicClientApplication | null,
    private readonly scopes: string[],
    private readonly logger: Logger,
    private readonly now: () => number = () => Date.now(),
    private readonly configError?: string,
  ) {}

  async init(): Promise<void> {
    if (!this.pca) return;
    this.account = await getCachedAccount(this.pca);
    this.state = this.account ? "token_expired" : "uninitialized";
  }

  status() {
    return {
      state: this.state,
      ...(this.configError && { configError: this.configError }),
      ...(this.token && { tokenExpiresOn: this.token.expiresOn.toISOString() }),
    };
  }

  accountId(): string {
    return this.account?.homeAccountId ?? "anonymous";
  }

  async getAccessToken(): Promise<string> {
    if (!this.pca) throw new AppError("config_error", this.configError ?? "Set MS_TODO_CLIENT_ID.");
    if (!this.account) throw new AppError("auth_required", LOGIN_HINT);
    if (this.token && this.token.expiresOn.getTime() - 60_000 > this.now()) return this.token.value;
    this.refreshing ??= this.refresh(this.pca, this.account).finally(() => { this.refreshing = null; });
    return this.refreshing;
  }

  private async refresh(pca: PublicClientApplication, account: AccountInfo): Promise<string> {
    try {
      const result = await pca.acquireTokenSilent({ account, scopes: this.scopes });
      validateGrant(result, this.scopes);
      if (!result.expiresOn || result.expiresOn.getTime() <= this.now() || result.account?.homeAccountId !== account.homeAccountId) {
        throw new Error("Invalid token response.");
      }
      this.token = { value: result.accessToken, expiresOn: result.expiresOn };
      this.state = "authenticated";
      return this.token.value;
    } catch {
      this.token = null;
      this.state = "refresh_failed";
      this.logger.warn("Microsoft authentication requires attention");
      throw new AppError("auth_required", LOGIN_HINT);
    }
  }
}
