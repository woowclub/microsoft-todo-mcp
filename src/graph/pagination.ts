import { pageUrl } from "./destination";
import { AppError } from "./errors";

/** OData collection response. */
export interface Page<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export interface PagedResult<T> {
  items: T[];
  hasMore: boolean;
  /** Opaque cursor (base64url of the next `@odata.nextLink`); pass it back to continue. */
  cursor?: string;
}

function encodeCursor(nextLink: string): string {
  return Buffer.from(nextLink, "utf-8").toString("base64url");
}

function decodeCursor(cursor: string): string {
  if (typeof cursor !== "string" || cursor.length > 10923 || !/^[A-Za-z0-9_-]+$/.test(cursor)) {
    throw new AppError("validation_error", "Invalid pagination cursor.");
  }
  const link = Buffer.from(cursor, "base64url").toString("utf8");
  if (encodeCursor(link) !== cursor) throw new AppError("validation_error", "Invalid pagination cursor.");
  return link;
}

const MAX_PAGES = 32;
const MAX_TOTAL_BYTES = 8 * 1024 * 1024;

/**
 * Follow `@odata.nextLink` until we have at least `softCap` items (page-granular — we never
 * split a page, so the result may slightly exceed the cap) or run out of pages. If stopped
 * by the cap with more to come, returns a resumable cursor.
 */
export async function collectPages<T>(
  fetchFirst: () => Promise<Page<T>>,
  fetchNext: (absoluteUrl: string) => Promise<Page<T>>,
  softCap: number,
  collection: string,
  startCursor?: string,
): Promise<PagedResult<T>> {
  const items: T[] = [];
  let url = startCursor === undefined ? undefined : pageUrl(decodeCursor(startCursor), collection);
  const visited = new Set<string>();
  let bytes = 0;

  for (let pages = 1; ; pages++) {
    if (url && visited.has(url)) throw new AppError("graph_error", "Microsoft Graph returned a repeated pagination link.");
    if (url) visited.add(url);
    const page: Page<T> = url ? await fetchNext(url) : await fetchFirst();
    if (!page || !Array.isArray(page.value) || page.value.length > 1000) {
      throw new AppError("graph_error", "Microsoft Graph returned an invalid collection.");
    }
    bytes += Buffer.byteLength(JSON.stringify(page.value));
    if (bytes > MAX_TOTAL_BYTES) throw new AppError("graph_error", "Microsoft Graph collection exceeded the size limit.");
    items.push(...page.value);
    const next = page["@odata.nextLink"];
    if (next === undefined) return { items, hasMore: false };
    url = pageUrl(next, collection);
    if (items.length >= softCap) return { items, hasMore: true, cursor: encodeCursor(url) };
    if (pages >= MAX_PAGES) throw new AppError("graph_error", `Pagination did not terminate after ${MAX_PAGES} pages.`);
  }
}
