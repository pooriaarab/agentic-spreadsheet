import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { ctxFromEnv } from "./env.js";
import { message, OPS, type Ctx } from "./ops.js";
import { VERSION } from "./version.js";

const text = (value: unknown) => [{ type: "text" as const, text: JSON.stringify(value) }];

// The server reads SHEET_AS and SHEET_KIND once. A write tool fails until both are set.
export function createMcpServer(ctx: Ctx = ctxFromEnv(process.env)): McpServer {
  const server = new McpServer({ name: "agentic-spreadsheet", version: VERSION });
  for (const op of OPS) {
    server.registerTool(
      `sheet_${op.name.replaceAll("-", "_")}`,
      { description: op.description, inputSchema: op.shape, annotations: { readOnlyHint: op.readOnly === true } },
      async (args: Record<string, unknown>) => {
        try {
          return { content: text(await op.run(args, ctx)) };
        } catch (error) {
          return { isError: true, content: text({ error: message(error) }) };
        }
      },
    );
  }
  return server;
}

export async function startMcp(): Promise<void> {
  await createMcpServer().connect(new StdioServerTransport());
}
