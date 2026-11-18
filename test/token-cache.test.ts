import { chmod, link, lstat, mkdir, mkdtemp, readFile, realpath, rm, stat, symlink, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type { TokenCacheContext } from "@azure/msal-node";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { clearCache, fileCachePlugin } from "../src/auth/token-cache";

const context = (value = "fake-cache"): TokenCacheContext => ({
  cacheHasChanged: true,
  tokenCache: { serialize: () => value, deserialize: () => {} },
} as unknown as TokenCacheContext);

describe("private cache persistence", () => {
  let dir: string;
  let path: string;
  beforeEach(async () => {
    dir = await mkdtemp(join(await realpath("/tmp"), "todo-cache-"));
    path = join(dir, "private", "write.json");
  });
  afterEach(async () => { await rm(dir, { force: true, recursive: true }); });

  it("creates an owner-only cache and replaces it atomically", async () => {
    const plugin = fileCachePlugin(path);
    for (const value of ["first", "second"]) {
      const ctx = context(value);
      await plugin.beforeCacheAccess(ctx);
      await plugin.afterCacheAccess(ctx);
      expect(await readFile(path, "utf8")).toBe(value);
    }
    expect((await lstat(join(dir, "private"))).mode & 0o777).toBe(0o700);
    expect((await stat(path)).mode & 0o777).toBe(0o600);
    await clearCache(path);
    await expect(plugin.beforeCacheAccess(context())).rejects.toThrow();
  });

  it("rejects file and directory symlinks, hardlinks, and unsafe modes", async () => {
    const parent = join(dir, "private");
    const target = join(dir, "other");
    await writeFile(target, "fixture", { mode: 0o600 });
    await mkdir(parent, { mode: 0o700 });
    for (const make of [symlink, link]) {
      await make(target, path);
      await expect(fileCachePlugin(path).beforeCacheAccess(context())).rejects.toThrow();
      await expect(clearCache(path)).rejects.toThrow();
      expect(await readFile(target, "utf8")).toBe("fixture");
      await rm(path);
    }
    await writeFile(path, "fixture", { mode: 0o600 });
    await chmod(path, 0o644);
    await expect(fileCachePlugin(path).beforeCacheAccess(context())).rejects.toThrow();
    await chmod(parent, 0o755);
    await expect(fileCachePlugin(path).beforeCacheAccess(context())).rejects.toThrow();
    await symlink(parent, join(dir, "alias"));
    await expect(fileCachePlugin(join(dir, "alias", "write.json")).beforeCacheAccess(context())).rejects.toThrow();
  });

  it("detects conflicting writes and prevents a pending refresh from restoring a cleared cache", async () => {
    const first = fileCachePlugin(path);
    const second = fileCachePlugin(path);
    const a = context("winner");
    const b = context("stale");
    await first.beforeCacheAccess(a);
    await second.beforeCacheAccess(b);
    await first.afterCacheAccess(a);
    await expect(second.afterCacheAccess(b)).rejects.toThrow();
    expect(await readFile(path, "utf8")).toBe("winner");
    await first.beforeCacheAccess(a);
    await clearCache(path);
    await expect(first.afterCacheAccess(a)).rejects.toThrow();
    await expect(stat(path)).rejects.toMatchObject({ code: "ENOENT" });
  });
});
