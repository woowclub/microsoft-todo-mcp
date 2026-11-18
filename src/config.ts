import { AppError } from "./graph/errors";
import { homedir } from "node:os";
import { join } from "node:path";

export const PERSONAL_AUTHORITY = "https://login.microsoftonline.com/consumers";
export const PERSONAL_TENANT = "9188040d-6c67-4c5b-b112-36a304b66dad";
export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const UUID = /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i;

export interface AppConfig {
  clientId: string;
  authority: string;
  tokenCachePath: string;
  /** Request only `Tasks.Read` (true read-only capability; requires re-login). */
  scopeReadonly: boolean;
  /** Runtime guard: hide write + destructive tools. */
  readonly: boolean;
  /** Runtime guard: hide destructive (delete) tools only. */
  noDestructive: boolean;
  /** Soft cap on items returned by list tools (page-granular). */
  maxResults: number;
  graphBase: string;
  logLevel: string;
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const clientId = env.MS_TODO_CLIENT_ID?.toLowerCase() ?? "";
  if (clientId && !UUID.test(clientId)) throw new AppError("config_error", "MS_TODO_CLIENT_ID must be an application UUID.");
  if ((env.MS_TODO_AUTHORITY && env.MS_TODO_AUTHORITY !== PERSONAL_AUTHORITY) ||
      (env.MS_TODO_GRAPH_BASE && env.MS_TODO_GRAPH_BASE !== GRAPH_BASE) || env.MS_TODO_TOKEN_CACHE) {
    throw new AppError("config_error", "This fork uses fixed Microsoft endpoints and its own private cache location.");
  }
  const scopeReadonly = envBool(env.MS_TODO_SCOPE_READONLY);
  return {
    clientId,
    authority: PERSONAL_AUTHORITY,
    tokenCachePath: join(homedir(), ".config", "microsoft-todo-mcp", clientId || "unconfigured", scopeReadonly ? "read.json" : "write.json"),
    scopeReadonly,
    readonly: scopeReadonly || envBool(env.MS_TODO_READONLY),
    noDestructive: env.MS_TODO_ALLOW_DESTRUCTIVE !== "1" || envBool(env.MS_TODO_NO_DESTRUCTIVE),
    maxResults: clampInt(env.MS_TODO_MAX_RESULTS, 200, 1, 1000),
    graphBase: GRAPH_BASE,
    logLevel: env.LOG_LEVEL ?? "info",
  };
}

export function scopesFor(config: AppConfig): string[] {
  // MSAL adds `offline_access` automatically for the refresh token.
  return config.scopeReadonly ? ["Tasks.Read"] : ["Tasks.ReadWrite"];
}

function envBool(v: string | undefined): boolean {
  return v === "1" || v?.toLowerCase() === "true";
}

function clampInt(v: string | undefined, fallback: number, min: number, max: number): number {
  const n = v ? Number.parseInt(v, 10) : Number.NaN;
  if (!Number.isFinite(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}

/** Only the configured To Do permission and MSAL's standard OIDC/refresh scopes. */
export function hasExpectedScopes(granted: string[], requested: string[]): boolean {
  const taskScope = requested[0]?.toLowerCase();
  if (requested.length !== 1 || (taskScope !== "tasks.readwrite" && taskScope !== "tasks.read")) return false;
  const normalized = granted.map((value) => value.toLowerCase().replace(/^https:\/\/graph\.microsoft\.com\//, ""));
  const allowed = new Set([taskScope, "openid", "profile", "offline_access"]);
  return normalized.includes(taskScope) && normalized.every((scope) => allowed.has(scope));
}
