import { constants } from "node:fs";
import { lstat, mkdir, open, rename, rmdir, unlink } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { dirname, isAbsolute, join, normalize, parse, sep } from "node:path";
import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import { AppError } from "../graph/errors";

const MAX_CACHE_BYTES = 4 * 1024 * 1024;
const unsafe = () => new AppError("config_error", "Unsafe or unreadable token cache. Check cache ownership and permissions.");
const missing = (err: unknown) => (err as NodeJS.ErrnoException).code === "ENOENT";

// Check every ancestor before descending or creating a directory. Root-owned sticky
// temporary directories are safe ancestors of user-owned private directories.
async function privateParent(path: string, create: boolean): Promise<boolean> {
  const uid = process.getuid?.();
  if (uid === undefined || !isAbsolute(path) || path !== normalize(path)) throw unsafe();
  const parent = dirname(path);
  let dir = parse(parent).root;
  if (dir === parent) throw unsafe();
  for (const segment of parent.slice(dir.length).split(sep)) {
    dir = join(dir, segment);
    if (create) {
      try { await mkdir(dir, { mode: 0o700 }); }
      catch (err) { if ((err as NodeJS.ErrnoException).code !== "EEXIST") throw unsafe(); }
    }
    let st;
    try { st = await lstat(dir); }
    catch (err) { if (!create && missing(err)) return false; throw unsafe(); }
    const trustedStickyDir = st.uid === 0 && (st.mode & 0o1000) !== 0;
    if (!st.isDirectory() || (st.uid !== uid && st.uid !== 0) ||
        ((st.mode & 0o022) !== 0 && !trustedStickyDir) ||
        (dir === parent && (st.uid !== uid || (st.mode & 0o077) !== 0))) throw unsafe();
  }
  return true;
}

async function readCache(path: string): Promise<string | null> {
  let file;
  try { file = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK); }
  catch (err) { if (missing(err)) return null; throw unsafe(); }
  try {
    const stat = await file.stat();
    if (!stat.isFile() || stat.nlink !== 1 || stat.uid !== process.getuid?.() ||
        (stat.mode & 0o777) !== 0o600 || stat.size > MAX_CACHE_BYTES) throw unsafe();
    return await file.readFile("utf8");
  } finally { await file.close(); }
}

// Serializes the short compare-and-replace step across processes. A crash here leaves
// a lock directory, which requires owner cleanup with all clients stopped.
async function withWriteLock(path: string, action: () => Promise<void>): Promise<void> {
  const lock = path + ".lock";
  try { await mkdir(lock, { mode: 0o700 }); }
  catch { throw new AppError("config_error", "Token cache is busy or locked. Retry after stopping other instances."); }
  try { await action(); }
  finally { await rmdir(lock); }
}

export function fileCachePlugin(path: string): ICachePlugin {
  const snapshots = new WeakMap<TokenCacheContext, string | null>();
  let previouslyPresent = false;
  return {
    async beforeCacheAccess(ctx) {
      const data = await privateParent(path, false) ? await readCache(path) : null;
      if (previouslyPresent && data === null) throw new AppError("auth_required", "Token cache cleared. Restart the server.");
      if (data !== null) {
        try { ctx.tokenCache.deserialize(data); }
        catch { throw unsafe(); }
        previouslyPresent = true;
      }
      snapshots.set(ctx, data);
    },
    async afterCacheAccess(ctx) {
      if (!ctx.cacheHasChanged) return;
      const snapshot = snapshots.get(ctx);
      if (snapshot === undefined) throw unsafe();
      await privateParent(path, true);
      await withWriteLock(path, async () => {
        if (await readCache(path) !== snapshot) throw new AppError("auth_required", "Token cache changed concurrently. Restart the server.");
        const data = ctx.tokenCache.serialize();
        if (Buffer.byteLength(data) > MAX_CACHE_BYTES) throw unsafe();
        const temporary = join(dirname(path), randomUUID() + ".tmp");
        const file = await open(temporary, "wx", 0o600);
        try {
          try { await file.writeFile(data, "utf8"); await file.sync(); }
          finally { await file.close(); }
          await rename(temporary, path);
          previouslyPresent = true;
        } finally {
          await unlink(temporary).catch((err) => { if (!missing(err)) throw err; });
        }
      });
    },
  };
}

export async function clearCache(path: string): Promise<void> {
  if (!await privateParent(path, false)) return;
  await withWriteLock(path, async () => {
    if (await readCache(path) !== null) await unlink(path);
  });
}
