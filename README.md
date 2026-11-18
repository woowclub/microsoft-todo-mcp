# Microsoft To Do MCP for personal macOS use

This fork runs over stdio. Codex can list and edit To Do data using a dedicated personal-account app registration. Task text returned by a tool enters the Codex conversation and can reach its configured model. Choose the lists and tasks you ask Codex to read accordingly.

## Build the reviewed checkout

Use a supported Node.js 24 LTS installation and pnpm 12.21.0. Run these commands from the repository root.

```sh
node --version
pnpm --version
pnpm install --frozen-lockfile --ignore-scripts
pnpm audit --prod
pnpm audit
pnpm typecheck
pnpm lint
pnpm test
pnpm build
node dist/index.js --version
```

Review any advisory before continuing. The review records one open development-tool advisory. Keep the checkout and its node_modules in a directory controlled by your macOS account. Codex launches Node directly from that checkout.

## Connect and verify

1. Complete the dedicated personal-account registration. [Registration instructions](docs/ENTRA-APP-SETUP.md)
2. In a private terminal, replace the placeholder below with the **Application (client) ID**. Register the executable once. The ID is public. The commands record absolute paths, so rerun registration if Node or the checkout moves.

   ```sh
   export MS_TODO_CLIENT_ID='your-application-client-id'
   codex mcp add microsoft-todo --env "MS_TODO_CLIENT_ID=$MS_TODO_CLIENT_ID" -- "$(command -v node)" "$PWD/dist/index.js" serve --no-destructive
   codex mcp get microsoft-todo
   ```

3. In the same terminal, run `node dist/index.js login`. Open Microsoft's device sign-in page yourself, enter the code from your terminal, select your personal account and inspect the app name and permissions. Enter credentials only on Microsoft's site. Keep the code out of chats, screenshots and logs.
4. Run `node dist/index.js whoami` and confirm the intended account. Start a fresh Codex session. Check the advertised tools and use `auth_status`. List your task lists, then create a task with a distinctive test title in a list you own. Read it back, complete it, reopen it and check it in Microsoft's To Do app. Remove it manually there if desired.

The default exposes list creation and renaming, task edits and checklist edits. Deletes require an explicit local opt-in. To enable later, replace the Codex registration with one that sets `--env MS_TODO_ALLOW_DESTRUCTIVE=1`, omits `--no-destructive`, and restart Codex. This enables list, task and checklist deletion. Runtime `--readonly` also disables writes. For a lower-privilege token use a dedicated registration with delegated Tasks.Read and set MS_TODO_SCOPE_READONLY=1 for both CLI and server.

## Stop, revoke or switch accounts

Close all Codex sessions using this server. With the original client ID and scope mode set, run `node dist/index.js logout`. Repeat for any other client ID or scope mode used. Remove the app's permissions on Microsoft's personal account consent page. Local logout removes a cache, while remote revocation withdraws consent. An issued access token may remain valid until expiry. To switch accounts, complete logout first, then repeat the terminal login and account check. To unregister the tool, run `codex mcp remove microsoft-todo`. [Microsoft account app permissions](https://account.live.com/consent/Manage)

## Update

Stop the server, review the incoming source, scripts, CI and lockfile diff against your reviewed revision, then update the local checkout deliberately. Repeat the build and audit commands above, inspect tool registration and restart Codex. The server uses local build output and local dependencies on every launch. [Security review](SECURITY_REVIEW.md)
