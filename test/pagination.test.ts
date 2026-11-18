import { describe, expect, it } from "vitest";
import * as fc from "fast-check";
import { collectPages, type Page } from "../src/graph/pagination";

const collection = "/me/todo/lists";
const base = "https://graph.microsoft.com/v1.0" + collection;
const next = (offset: number) => base + "?$skip=" + offset;
const cursor = (link: string) => Buffer.from(link).toString("base64url");

describe("collectPages", () => {
  function splitIntoPages<T>(items: T[], pageSize: number): Page<T>[] {
    const pages: Page<T>[] = [];
    for (let i = 0; i < items.length; i += pageSize) {
      const hasNext = i + pageSize < items.length;
      pages.push({ value: items.slice(i, i + pageSize), ...(hasNext ? { "@odata.nextLink": next(i + pageSize) } : {}) });
    }
    if (pages.length === 0) pages.push({ value: [] });
    return pages;
  }

  it("concatenates every item in order with no duplicates when the cap is high", async () => {
    await fc.assert(
      fc.asyncProperty(fc.array(fc.integer(), { maxLength: 30 }), fc.integer({ min: 1, max: 40 }), async (items, pageSize) => {
        const pages = splitIntoPages(items, pageSize);
        let idx = 0;
        const result = await collectPages<number>(
          async () => pages[0]!,
          async () => pages[++idx]!,
          items.length + 1,
          collection,
        );
        expect(result.items).toEqual(items);
        expect(result.hasMore).toBe(false);
        expect(result.cursor).toBeUndefined();
      }),
    );
  });

  it("stops at the soft cap (page-granular) and returns a resumable cursor", async () => {
    const pages: Page<number>[] = [
      { value: [1, 2, 3], "@odata.nextLink": next(3) },
      { value: [4, 5, 6], "@odata.nextLink": next(6) },
      { value: [7, 8] },
    ];
    let idx = 0;
    const result = await collectPages<number>(
      async () => pages[0]!,
      async (url) => {
        expect(url).toBe(next(3));
        return pages[++idx]!;
      },
      2,
      collection,
    );
    expect(result.items).toEqual([1, 2, 3]);
    expect(result.hasMore).toBe(true);
    expect(result.cursor).toBe(cursor(next(3)));
    expect(idx).toBe(0); // fetchNext never called: cap reached after the first page
  });

  it("resumes from a cursor", async () => {
    const pages: Record<string, Page<number>> = {
      [next(3)]: { value: [4, 5, 6] },
    };
    const result = await collectPages<number>(
      async () => {
        throw new Error("fetchFirst must not be called when a cursor is supplied");
      },
      async (url) => pages[url]!,
      100,
      collection,
      cursor(next(3)),
    );
    expect(result.items).toEqual([4, 5, 6]);
    expect(result.hasMore).toBe(false);
  });

  it("throws rather than looping forever on a non-terminating nextLink chain", async () => {
    const endless = async () => ({ value: [] as number[], "@odata.nextLink": next(1) });
    await expect(collectPages<number>(endless, endless, 10, collection)).rejects.toMatchObject({ code: "graph_error" });
  });
});

describe("untrusted pagination input", () => {
  const rejected = [
    "https://untrusted.invalid/v1.0/me/todo/lists",
    "https://graph.microsoft.com.invalid/v1.0/me/todo/lists",
    "https://graph.microsoft.com@untrusted.invalid/v1.0/me/todo/lists",
    "https://user@graph.microsoft.com/v1.0/me/todo/lists",
    "https://graph.microsoft.com:444/v1.0/me/todo/lists",
    "http://graph.microsoft.com/v1.0/me/todo/lists",
    "file:///v1.0/me/todo/lists",
    "https://graph.microsoft.com/v1.0/me/messages",
    base + "/A/tasks",
    base + "/%2e%2e/lists",
    base + "/%2F",
    base + "?$expand=extensions",
    base + "?$top=1&$TOP=2",
    base + "#fragment",
    "/me/todo/lists",
  ];
  it.each(rejected)("rejects cursors and Graph links before following or returning them (%s)", async (link) => {
    let calls = 0;
    const fetchNext = async () => { calls++; return { value: [] }; };
    const first = async () => ({ value: [1], "@odata.nextLink": link });
    for (const cap of [1, 10]) {
      await expect(collectPages(first, fetchNext, cap, collection)).rejects.toMatchObject({ code: "validation_error" });
    }
    await expect(collectPages(first, fetchNext, 1, collection, cursor(link))).rejects.toMatchObject({ code: "validation_error" });
    expect(calls).toBe(0);
  });

  it.each(["", "%", "AB", cursor(next(2)) + "=", "_w", "a".repeat(10924)])("rejects malformed cursors", async (invalid) => {
    const unexpected = async () => { throw new Error("unexpected HTTP request"); };
    await expect(collectPages(unexpected, unexpected, 1, collection, invalid)).rejects.toMatchObject({ code: "validation_error" });
  });
});
