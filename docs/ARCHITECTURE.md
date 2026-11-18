# Architecture

The CLI selects terminal authentication, cache management, account inspection or a stdio MCP server. The server registers allowed operations and validates each tool call. Task and list adapters translate these operations to Microsoft Graph. The Graph client enforces destination and method policy before sending a bearer-authenticated request. Pagination validates continuation links against the requested collection.

MSAL manages device authentication and silent refresh. The credential cache is partitioned by client ID and scope mode. The server's in-memory create-operation cache is bounded and scoped to the signed-in account. Runtime dependencies are MSAL Node, the MCP SDK and Zod. The pnpm build uses esbuild and leaves dependencies external. [Security assumptions and test coverage](../SECURITY_REVIEW.md)
