import { describe, expect, it, vi } from "vitest";
import { GraphClient } from "../src/graph/client";
import { TaskListsApi } from "../src/todo/lists";
import { TasksApi } from "../src/todo/tasks";
import { createLogger } from "../src/lib/logger";
import type { AppConfig } from "../src/config";

const config = { noDestructive: false } as AppConfig;
const silent = createLogger("silent");
const tokens = { getAccessToken: async () => "test-only-internal-access-credential", accountId: () => "acct" };
const ids = { uuid: () => "00000000-0000-0000-0000-000000000000" };
const noSleep = async () => {};

/** A fetch that never resolves on its own — it only settles when the AbortSignal fires. */
function hangingFetch(onCall?: () => void): typeof fetch {
  return ((_url: string, opts?: { signal?: AbortSignal }) => {
    onCall?.();
    return new Promise<Response>((_resolve, reject) => {
      opts?.signal?.addEventListener("abort", () => reject(new DOMException("The operation timed out.", "TimeoutError")));
    });
  }) as unknown as typeof fetch;
}

function jsonResponse(status: number, body: unknown, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json", ...headers } });
}

describe("GraphClient", () => {
  it("times out a hung request instead of hanging forever", async () => {
    const client = new GraphClient({ config, tokens, logger: silent, ids, requestTimeoutMs: 20, fetchImpl: hangingFetch(), sleep: noSleep });
    await expect(client.request("GET", "/me/todo/lists")).rejects.toMatchObject({ code: "graph_unavailable" });
  });

  it("does not auto-retry a POST that may have landed", async () => {
    let calls = 0;
    const client = new GraphClient({ config, tokens, logger: silent, ids, requestTimeoutMs: 20, fetchImpl: hangingFetch(() => { calls += 1; }), sleep: noSleep });
    await expect(client.request("POST", "/me/todo/lists", { body: { displayName: "t" } })).rejects.toMatchObject({ code: "graph_unavailable" });
    expect(calls).toBe(1);
  });

  it("retries a GET on a 503, then succeeds", async () => {
    let calls = 0;
    const fetchImpl = (async () => {
      calls += 1;
      return calls < 2 ? jsonResponse(503, { error: { code: "serviceUnavailable" } }) : jsonResponse(200, { value: [] });
    }) as unknown as typeof fetch;
    const client = new GraphClient({ config, tokens, logger: silent, ids, fetchImpl, sleep: noSleep });
    const res = await client.request<{ value: unknown[] }>("GET", "/me/todo/lists");
    expect(res.value).toEqual([]);
    expect(calls).toBe(2);
  });

  it("maps a 404 to not_found and carries only the locally generated correlation id", async () => {
    const fetchImpl = (async () => jsonResponse(404, { error: { code: "ItemNotFound", message: "not found" } }, { "request-id": "req-abc-123" })) as unknown as typeof fetch;
    const client = new GraphClient({ config, tokens, logger: silent, ids, fetchImpl, sleep: noSleep });
    await expect(client.request("GET", "/me/todo/lists/nope")).rejects.toMatchObject({ code: "not_found", clientRequestId: ids.uuid() });
  });

  it("returns undefined for a 204 (no body)", async () => {
    const fetchImpl = vi.fn<typeof fetch>()
      .mockResolvedValueOnce(jsonResponse(200, { id: "x", isOwner: true }))
      .mockResolvedValueOnce(new Response(null, { status: 204 }));
    const client = new GraphClient({ config, tokens, logger: silent, ids, fetchImpl, sleep: noSleep });
    expect(await client.request("DELETE", "/me/todo/lists/x/tasks/y")).toBeUndefined();
  });
});

describe("authenticated HTTP boundary", () => {
  const fakeToken = "FAKE-test-only-credential";
  const mockTokens = { ...tokens, getAccessToken: vi.fn(async () => fakeToken) };
  it("rejects external destinations before obtaining credentials or invoking fetch", async () => {
    const fetchImpl = vi.fn<typeof fetch>();
    mockTokens.getAccessToken.mockClear();
    const client = new GraphClient({ config, tokens: mockTokens, logger: silent, ids, fetchImpl });
    for (const url of ["https://other.invalid/", "//other.invalid", "/me/mailFolders", "/me/todo/lists/../lists"]) {
      await expect(client.request("GET", url)).rejects.toMatchObject({ code: "validation_error" });
    }
    expect(fetchImpl).not.toHaveBeenCalled();
    expect(mockTokens.getAccessToken).not.toHaveBeenCalled();
  });

  it("keeps credentials inside Graph when list tools receive a hostile cursor or nextLink", async () => {
    const external = "https://untrusted.invalid/v1.0/me/todo/lists";
    const fetchImpl = vi.fn<typeof fetch>().mockImplementation(async () => jsonResponse(200, {
      value: [{ id: "L", displayName: "sample" }], "@odata.nextLink": external,
    }));
    const client = new GraphClient({ config, tokens: mockTokens, logger: silent, ids, fetchImpl });
    const lists = new TaskListsApi(client, 1);
    const tasks = new TasksApi(client, 1);
    const cursor = Buffer.from(external).toString("base64url");
    await expect(lists.list({ cursor })).rejects.toMatchObject({ code: "validation_error" });
    await expect(tasks.list("L", { cursor })).rejects.toMatchObject({ code: "validation_error" });
    expect(fetchImpl).not.toHaveBeenCalled();
    await expect(lists.list()).rejects.toMatchObject({ code: "validation_error" });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    const [url, init] = fetchImpl.mock.calls[0]!;
    expect(String(url)).toBe("https://graph.microsoft.com/v1.0/me/todo/lists?%24top=100");
    expect(new Headers(init?.headers).get("authorization")).toBe("Bearer " + fakeToken);
  });

  it("withholds a successful response reflecting the current bearer in a legitimate content field", async () => {
    const fetchImpl = vi.fn<typeof fetch>().mockResolvedValue(jsonResponse(200, {
      value: [{ id: "L", isOwner: true, displayName: "Sample " + fakeToken }],
    }));
    const client = new GraphClient({ config, tokens: mockTokens, logger: silent, ids, fetchImpl });
    const error = await new TaskListsApi(client, 1).list().catch((err: Error) => err);
    expect(error).toMatchObject({ code: "graph_error" });
    expect(String(error)).not.toContain(fakeToken);
  });

  it("projects known task and linked-resource fields into MCP results", async () => {
    const fetchImpl = vi.fn<typeof fetch>().mockResolvedValueOnce(jsonResponse(200, { id: "L", isOwner: true })).mockResolvedValueOnce(jsonResponse(200, {
      id: "T", title: "fixture", access_token: fakeToken,
      body: { content: "text", contentType: "text", extra: fakeToken },
      linkedResources: [{ id: "R", displayName: "source", extra: fakeToken }],
    }));
    const client = new GraphClient({ config, tokens, logger: silent, ids, fetchImpl });
    const tasks = new TasksApi(client, 10);
    const result = await tasks.get("L", "T", { includeLinkedResources: true });
    expect(result.linkedResources).toEqual([{ id: "R", displayName: "source" }]);
    expect(JSON.stringify(result)).not.toContain(fakeToken);
  });

  it("pins the destination and disables redirect following on the actual authenticated request", async () => {
    const injected = { ...config, graphBase: "https://other.invalid" };
    const fetchImpl = vi.fn<typeof fetch>().mockResolvedValue(new Response(null, {
      status: 307, headers: { location: "https://other.invalid/redirect" },
    }));
    const client = new GraphClient({ config: injected, tokens: mockTokens, logger: silent, ids, fetchImpl });
    await expect(client.request("GET", "/me/todo/lists")).rejects.toMatchObject({ code: "graph_error" });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    const [url, init] = fetchImpl.mock.calls[0]!;
    expect(url).toBe("https://graph.microsoft.com/v1.0/me/todo/lists");
    expect(init?.redirect).toBe("manual");
    expect(new Headers(init?.headers).get("authorization")).toBe("Bearer " + fakeToken);
  });

  it("keeps HTTP error bodies, headers and transport exceptions out of errors and logs", async () => {
    const logs = vi.spyOn(process.stderr, "write").mockReturnValue(true);
    try {
      for (const fetchImpl of [
        vi.fn<typeof fetch>().mockRejectedValue(new Error(fakeToken)),
        vi.fn<typeof fetch>().mockResolvedValue(jsonResponse(404, { error: { message: fakeToken, code: fakeToken } },
          { "request-id": fakeToken, "client-request-id": fakeToken })),
      ]) {
        const client = new GraphClient({ config, tokens: mockTokens, logger: createLogger("debug"), ids, fetchImpl, sleep: noSleep });
        const error = await client.request("GET", "/me/todo/lists/" + fakeToken).catch((err: Error) => err);
        expect(error).toBeInstanceOf(Error);
        expect(JSON.stringify(error)).not.toContain(fakeToken);
        expect(String(error)).not.toContain(fakeToken);
      }
      expect(JSON.stringify(logs.mock.calls)).not.toContain(fakeToken);
    } finally {
      logs.mockRestore();
    }
  });

  it("rejects disallowed mutations before transport", async () => {
    const fetchImpl = vi.fn<typeof fetch>();
    for (const [method, guarded] of [
      ["DELETE", { ...config, noDestructive: true }],
      ["PATCH", { ...config, readonly: true }],
      ["POST", { ...config, scopeReadonly: true }],
    ] as const) {
      const client = new GraphClient({ config: guarded, tokens, logger: silent, ids, fetchImpl });
      await expect(client.request(method, "/me/todo/lists/test")).rejects.toMatchObject({ code: "permission_denied" });
    }
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("rejects wrong content types, malformed JSON, and oversized success responses", async () => {
    for (const response of [
      new Response("{}", { headers: { "content-type": "text/html" } }),
      new Response("invalid", { headers: { "content-type": "application/json" } }),
      jsonResponse(200, "x".repeat(4 * 1024 * 1024)),
    ]) {
      const client = new GraphClient({ config, tokens, logger: silent, ids, fetchImpl: async () => response });
      await expect(client.request("GET", "/me/todo/lists")).rejects.toMatchObject({ code: "graph_error" });
    }
  });
});

describe("owner-only To Do access", () => {
  it("omits shared-in lists and lists with missing ownership metadata", async () => {
    const fetchImpl = vi.fn<typeof fetch>().mockResolvedValue(jsonResponse(200, { value: [
      { id: "owned", displayName: "owned", isOwner: true },
      { id: "shared", displayName: "shared", isOwner: false },
      { id: "unknown", displayName: "unknown" },
    ] }));
    const client = new GraphClient({ config, tokens, logger: silent, ids, fetchImpl });
    const result = await new TaskListsApi(client, 100).list();
    expect(result.items.map((list) => list.id)).toEqual(["owned"]);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it.each([{ id: "L", isOwner: false }, { id: "L" }, { id: "different", isOwner: true }])(
    "rejects direct IDs and cursors without positive matching ownership",
    async (metadata) => {
      const fetchImpl = vi.fn<typeof fetch>().mockImplementation(async () => jsonResponse(200, metadata));
      const client = new GraphClient({ config, tokens, logger: silent, ids, fetchImpl });
      const list = "https://graph.microsoft.com/v1.0/me/todo/lists/L";
      const tasks = new TasksApi(client, 10);
      const cursor = Buffer.from(list + "/tasks?$skip=100").toString("base64url");
      const calls = [
        () => tasks.list("L", { cursor }),
        () => tasks.create({ listId: "L", title: "sample" }),
        () => tasks.listChecklist("L", "T"),
        () => tasks.remove("L", "T"),
      ];
      for (const operation of calls) await expect(operation()).rejects.toMatchObject({ code: "permission_denied" });
      expect(fetchImpl).toHaveBeenCalledTimes(calls.length);
      for (const [url, init] of fetchImpl.mock.calls) {
        expect(url).toBe(list + "?$select=id,isOwner");
        expect(init?.method).toBe("GET");
      }
    },
  );
});
