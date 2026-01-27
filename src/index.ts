#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerTools } from "./tools.js";

const server = new McpServer({
  name: "numbers-mcp-server",
  version: "1.0.0",
});

// Register all tools
registerTools(server);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // Use stderr for logging - stdout is reserved for JSON-RPC
  console.error("[numbers-mcp] Server running on stdio");
}

main().catch((err) => {
  console.error("[numbers-mcp] Fatal error:", err);
  process.exit(1);
});
