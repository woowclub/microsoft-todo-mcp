import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createToolServer } from "../src/server";
import { loadConfig } from "../src/config";
import { createLogger } from "../src/lib/logger";
import type { ToolContext } from "../src/tools/_shared";
import { describe, expect, it } from "vitest";
import { allTools } from "../src/tools/registry";

describe("tool registry", () => {
  it("exposes the expected tool surface", () => {
    expect(allTools).toHaveLength(16);
  });

  it("has unique tool names", () => {
    const names = allTools.map((t) => t.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it("gives every tool a description, a valid risk class, and an input shape", () => {
    for (const tool of allTools) {
      expect(tool.description.length).toBeGreaterThan(0);
      expect(["read", "write", "destructive"]).toContain(tool.risk);
      expect(typeof tool.inputShape).toBe("object");
    }
  });

  it("warns in the description of every destructive tool", () => {
    const destructive = allTools.filter((t) => t.risk === "destructive");
    expect(destructive.length).toBeGreaterThan(0);
    for (const tool of destructive) {
      expect(tool.description.toLowerCase()).toContain("cannot be undone");
    }
  });
});

describe("MCP protocol tool boundary", () => {
  async function inspect(env: NodeJS.ProcessEnv, check: (client: Client) => Promise<void>) {
    const config = loadConfig(env);
    const logger = createLogger("silent");
    // Omit the Graph adapters so an accidental dispatch fails the test.
    const server = createToolServer(config, { logger } as ToolContext);
    const client = new Client({ name: "offline-test", version: "1" });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await server.connect(serverTransport);
    try {
      await client.connect(clientTransport);
      await check(client);
    } finally {
      await client.close();
      await server.close();
    }
  }

  it("exposes only approved personal To Do operations by default and refuses hidden calls", async () => {
    await inspect({}, async client => {
      const { tools } = await client.listTools();
      expect(tools.map(t => t.name).sort()).toEqual([
        "add_checklist_item", "auth_status", "complete_task", "create_task", "create_task_list",
        "get_task", "list_checklist_items", "list_task_lists", "list_tasks", "sign_in",
        "update_checklist_item", "update_task", "update_task_list",
      ]);
      for (const name of ["delete_task", "delete_task_list", "delete_checklist_item"]) {
        const res = await client.callTool({ name, arguments: { listId: "list", taskId: "task", itemId: "item" } });
        expect(res.isError).toBe(true);
        expect(JSON.stringify(res.content)).toContain("Tool not available");
      }
      const unknown = await client.callTool({ name: "SYNTHETIC-SECRET", arguments: {} });
      expect(JSON.stringify(unknown.content)).not.toContain("SYNTHETIC-SECRET");
      expect(client.getInstructions()).toContain("untrusted user data");
    });
  });

  it("requires explicit delete opt-in and gives read-only mode precedence", async () => {
    await inspect({ MS_TODO_ALLOW_DESTRUCTIVE: "1" }, async client => {
      expect((await client.listTools()).tools.filter(t => t.name.startsWith("delete_"))).toHaveLength(3);
    });
    await inspect({ MS_TODO_ALLOW_DESTRUCTIVE: "1", MS_TODO_READONLY: "1" }, async client => {
      expect((await client.listTools()).tools.every(t => t.annotations?.readOnlyHint)).toBe(true);
    });
  });

  it("rejects malformed IDs, oversized arguments, unexpected fields and invalid dates before dispatch", async () => {
    await inspect({}, async client => {
      for (const args of [
        { listId: "..", title: "t" },
        { listId: "list", title: "x".repeat(1025) },
        { listId: "list", title: "x", "SYNTHETIC-SECRET": true },
        { listId: "list", title: "x", dueDateTime: "some day" },
      ]) {
        const result = await client.callTool({ name: "create_task", arguments: args });
        expect(result.isError).toBe(true);
        expect(JSON.stringify(result.content)).toContain("Invalid arguments");
        expect(JSON.stringify(result.content)).not.toContain("SYNTHETIC-SECRET");
      }
    });
  });
});
