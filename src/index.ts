#!/usr/bin/env node
import { loadConfig, scopesFor } from "./config";
import { runServer } from "./server";
import { acquireByDeviceCode, createPca, getCachedAccount } from "./auth/device-code";
import { clearCache } from "./auth/token-cache";
import { AppError, toAppError } from "./graph/errors";
import { VERSION } from "./version";

interface ParsedArgs {
  cmd: string;
  flags: Set<string>;
}

function parseArgs(argv: string[]): ParsedArgs {
  const flags = new Set<string>();
  const positional: string[] = [];
  for (const arg of argv) {
    if (arg.startsWith("--")) flags.add(arg.slice(2));
    else positional.push(arg);
  }
  return { cmd: positional[0] ?? "serve", flags };
}

/** Translate `serve` flags into config overrides (these take precedence over env vars). */
function flagOverrides(flags: Set<string>): Record<string, string> {
  const overrides: Record<string, string> = {};
  if (flags.has("readonly")) overrides.MS_TODO_READONLY = "1";
  if (flags.has("no-destructive")) overrides.MS_TODO_NO_DESTRUCTIVE = "1";
  if (flags.has("scope-readonly")) overrides.MS_TODO_SCOPE_READONLY = "1";
  return overrides;
}

const HELP = `microsoft-todo-mcp ${VERSION}
Model Context Protocol server for Microsoft To Do.

Usage:
  microsoft-todo-mcp [serve]      Start the MCP server on stdio (default)
  microsoft-todo-mcp login        Sign in via device code (one time)
  microsoft-todo-mcp logout       Clear the cached token
  microsoft-todo-mcp whoami       Show the signed-in account
  microsoft-todo-mcp --version    Print version

Flags (serve):
  --readonly         Disable all write and destructive tools
  --no-destructive   Disable destructive (delete) tools only
  --scope-readonly   Request only the Tasks.Read scope (requires re-login)

Env:
  MS_TODO_CLIENT_ID, MS_TODO_MAX_RESULTS, MS_TODO_ALLOW_DESTRUCTIVE, LOG_LEVEL

First-time setup: docs/ENTRA-APP-SETUP.md
`;

async function main(): Promise<void> {
  const { cmd, flags } = parseArgs(process.argv.slice(2));

  if (flags.has("version") || cmd === "version") {
    process.stdout.write(`${VERSION}\n`);
    return;
  }
  if (flags.has("help") || cmd === "help") {
    process.stdout.write(HELP);
    return;
  }

  const config = loadConfig({ ...process.env, ...flagOverrides(flags) });

  switch (cmd) {
    case "serve":
      await runServer(config);
      return;

    case "login": {
      if (!process.stdin.isTTY || !process.stdout.isTTY) throw new AppError("auth_required", "Run login in an interactive terminal.");
      const pca = createPca(config);
      if (await getCachedAccount(pca)) throw new AppError("auth_required", "Stop the server and run logout before starting a new login.");
      await acquireByDeviceCode(pca, scopesFor(config), (code, verificationUri) => {
        process.stdout.write(`Open ${verificationUri} and enter this code:\n${code}\nReview the account, app and permissions before consenting.\n`);
      });
      process.stdout.write("Signed in. Restart your MCP server.\n");
      return;
    }

    case "logout":
      if (!config.clientId) throw new AppError("config_error", "Set the original MS_TODO_CLIENT_ID before logging out.");
      await clearCache(config.tokenCachePath);
      process.stdout.write("Signed out (token cache cleared).\n");
      return;

    case "whoami": {
      const account = await getCachedAccount(createPca(config));
      if (!account) {
        process.stdout.write("Not signed in. Run `microsoft-todo-mcp login`.\n");
        return;
      }
      process.stdout.write(JSON.stringify({ username: account.username, tenantId: account.tenantId }) + "\n");
      return;
    }

    default:
      process.stderr.write(`Unknown command.\n\n${HELP}`);
      process.exitCode = 1;
  }
}

main().catch((err) => {
  const appError = toAppError(err);
  process.stderr.write(
    JSON.stringify({ level: "error", time: new Date().toISOString(), msg: "fatal", error: appError.toJSON() }) + "\n",
  );
  process.exitCode = 1;
});
