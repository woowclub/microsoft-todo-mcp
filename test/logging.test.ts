import { expect, it, vi } from "vitest";
import { createLogger } from "../src/lib/logger";
import { errorResult } from "../src/tools/_shared";

it("omits unexpected provider data, raw headers, paths and nested credentials from diagnostics", () => {
  const secret = "SYNTHETIC-CREDENTIAL";
  const write = vi.spyOn(process.stderr, "write").mockReturnValue(true);
  try {
    createLogger("debug").child({ op: secret }).warn("provider failed", {
      status: 503, error: secret, msal: secret, requestId: secret,
      unknown: { arbitrary: { nested: { data: { token: secret } } } },
    });
    const line = String(write.mock.calls[0]?.[0]);
    expect(line).toContain('"status":503');
    expect(line).not.toContain(secret);
    expect(JSON.stringify(errorResult(new Error(secret)))).not.toContain(secret);
  } finally {
    write.mockRestore();
  }
});
