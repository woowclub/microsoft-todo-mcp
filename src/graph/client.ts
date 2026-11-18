import { readBoundedText } from "../lib/http-body";
import { graphUrl } from "./destination";
import { AppError, mapGraphError, retryAfterSeconds } from "./errors";
import { GRAPH_BASE, type AppConfig } from "../config";
import type { IdGenerator } from "../lib/ids";
import type { Logger } from "../lib/logger";

/** Supplies an access token to the Graph client; throws an `auth_*` AppError when it can't. */
export interface TokenProvider {
  getAccessToken(): Promise<string>;
  /** Stable id of the signed-in account (for idempotency-key scoping). */
  accountId(): string;
}

export interface GraphClientDeps {
  config: AppConfig;
  tokens: TokenProvider;
  logger: Logger;
  ids: IdGenerator;
  /** Per-request timeout in ms (default 15s). Mostly here so tests can use a short one. */
  requestTimeoutMs?: number;
  /** Overrides for tests. */
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
}

interface RequestOptions {
  query?: Record<string, string | number | boolean | undefined>;
  body?: unknown;
  traceId?: string;
}

const RETRYABLE_STATUS = new Set([429, 503, 504]);
const MAX_RETRIES = 3;
const MAX_BACKOFF_MS = 8_000;
const DEFAULT_REQUEST_TIMEOUT_MS = 15_000;

export class GraphClient {
  private readonly fetchImpl: typeof fetch;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly requestTimeoutMs: number;

  constructor(private readonly deps: GraphClientDeps) {
    this.fetchImpl = deps.fetchImpl ?? fetch;
    this.sleep = deps.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.requestTimeoutMs = deps.requestTimeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS;
  }

  /** `pathOrUrl` is either a path relative to the Graph base or a full URL (e.g. an `@odata.nextLink`). */
  async request<T>(method: string, pathOrUrl: string, opts: RequestOptions = {}): Promise<T> {
    const url = this.buildUrl(pathOrUrl, opts.query);
    if (!["GET", "POST", "PATCH", "DELETE"].includes(method) ||
        (method !== "GET" && (this.deps.config.readonly || this.deps.config.scopeReadonly)) ||
        (method === "DELETE" && this.deps.config.noDestructive !== false)) {
      throw new AppError("permission_denied", "This operation is disabled.");
    }
    const listId = new URL(url).pathname.split("/")[5];
    if (listId) {
      // Tasks.ReadWrite also covers shared lists. Check Graph's ownership flag before using
      // a caller-supplied ID, including when the caller bypasses list discovery with a cursor.
      const ownership = await this.send<{ id?: unknown; isOwner?: unknown } | null>(
        "GET", graphUrl(GRAPH_BASE + "/me/todo/lists/" + listId + "?$select=id,isOwner").href, { traceId: opts.traceId },
      );
      if (ownership?.isOwner !== true || ownership.id !== decodeURIComponent(listId)) {
        throw new AppError("permission_denied", "Access requires a task list owned by the signed-in account.");
      }
    }
    return this.send<T>(method, url, opts);
  }

  private async send<T>(method: string, url: string, opts: RequestOptions): Promise<T> {
    const log = this.deps.logger.child({ traceId: opts.traceId, op: method });
    const canRetry = method === "GET";

    for (let attempt = 0; ; attempt++) {
      const clientRequestId = this.deps.ids.uuid();
      const token = await this.deps.tokens.getAccessToken();
      const startedAt = Date.now();

      let res: Response;
      let bodyText: string;
      try {
        res = await this.fetchImpl(url, {
          method,
          redirect: "manual",
          headers: {
            Authorization: `Bearer ${token}`,
            "client-request-id": clientRequestId,
            Accept: "application/json",
            ...(opts.body !== undefined ? { "Content-Type": "application/json" } : {}),
          },
          body: opts.body !== undefined ? JSON.stringify(opts.body) : undefined,
          signal: AbortSignal.timeout(this.requestTimeoutMs),
        });
        if (res.status >= 300 && res.status < 400) {
          await res.body?.cancel();
          throw new AppError("graph_error", "Microsoft Graph returned a redirect.");
        }
        if (!res.ok || res.status === 204) {
          await res.body?.cancel();
          bodyText = "";
        } else {
          const mediaType = res.headers.get("content-type")?.split(";")[0]?.trim().toLowerCase();
          if (mediaType !== "application/json") {
            await res.body?.cancel();
            throw new AppError("graph_error", "Microsoft Graph returned an unexpected response type.");
          }
          bodyText = await readJsonText(res);
        }
      } catch (err) {
        if (err instanceof AppError) throw err;
        const timedOut = err instanceof Error && (err.name === "TimeoutError" || err.name === "AbortError");
        if (canRetry && attempt < MAX_RETRIES) {
          const waitMs = this.backoffMs(attempt);
          log.warn(timedOut ? "request timed out — retrying" : "transport error — retrying", { clientRequestId, attempt, waitMs, outcome: "error" });
          await this.sleep(waitMs);
          continue;
        }
        throw new AppError(
          "graph_unavailable",
          timedOut
            ? `Microsoft Graph did not respond within ${Math.round(this.requestTimeoutMs / 1000)}s.`
            : "Network error talking to Microsoft Graph.",
          { clientRequestId },
        );
      }

      const durationMs = Date.now() - startedAt;

      if (res.ok) {
        log.debug("graph ok", { clientRequestId, status: res.status, durationMs, outcome: "ok" });
        if (res.status === 204) return undefined as T;
        try {
          return JSON.parse(bodyText, (key: string, value: unknown) => {
            if (key.includes(token) || (typeof value === "string" && value.includes(token))) {
              throw new AppError("graph_error", "Microsoft Graph returned sensitive authentication data.");
            }
            return value;
          }) as T;
        } catch (err) {
          if (err instanceof AppError) throw err;
          throw new AppError("graph_error", "Microsoft Graph returned a response that was not valid JSON.", { status: res.status, clientRequestId });
        }
      }

      if (canRetry && RETRYABLE_STATUS.has(res.status) && attempt < MAX_RETRIES) {
        const retryAfter = retryAfterSeconds(res.headers);
        // Honour Retry-After, but cap it — a tool call shouldn't block for minutes.
        const waitMs = retryAfter === undefined ? this.backoffMs(attempt) : retryAfter * 1000;
        log.warn("retryable graph error", { clientRequestId, status: res.status, attempt, waitMs, durationMs, outcome: "error" });
        await this.sleep(waitMs);
        continue;
      }

      const error = mapGraphError(res.status, clientRequestId, res.headers);
      log.error("graph error", {
        clientRequestId: error.clientRequestId ?? clientRequestId,
        status: res.status,
        durationMs,
        outcome: "error",
        errorCode: error.code,
      });
      throw error;
    }
  }

  private buildUrl(pathOrUrl: string, query?: RequestOptions["query"]): string {
    const url = graphUrl(pathOrUrl.startsWith("/") ? GRAPH_BASE + pathOrUrl : pathOrUrl);
    if (query) {
      for (const [key, value] of Object.entries(query)) {
        if (value !== undefined) url.searchParams.set(key, String(value));
      }
    }
    return graphUrl(url.href).href;
  }

  private backoffMs(attempt: number): number {
    const ceiling = Math.min(MAX_BACKOFF_MS, 500 * 2 ** attempt);
    // Full jitter in [ceiling/2, ceiling].
    return Math.floor(ceiling / 2 + Math.random() * (ceiling / 2));
  }
}

async function readJsonText(response: Response): Promise<string> {
  return readBoundedText(response, 4 * 1024 * 1024, new AppError("graph_error", "Microsoft Graph response was empty or exceeded the size limit."));
}
