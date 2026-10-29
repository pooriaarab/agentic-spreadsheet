import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { ctxFromEnv } from "./env.js";
import { message, OPS, type Ctx } from "./ops.js";
import { remoteFromEnv, runShareOp, SHARE_OPS, type Remote } from "./remote.js";
import { VERSION } from "./version.js";

const text = (value: unknown) => [{ type: "text" as const, text: JSON.stringify(value) }];

// The server reads SHEET_AS and SHEET_KIND once. A write tool fails until both are set.
// With a Worker (`remote`), the server also has the sharing tools: roles live on the Worker, and a local store has none.
export function createMcpServer(ctx: Ctx = ctxFromEnv(process.env), remote?: Remote): McpServer {
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
  for (const op of remote === undefined ? [] : SHARE_OPS) {
    server.registerTool(
      `sheet_${op.name}`,
      { description: op.description, inputSchema: op.shape, annotations: { readOnlyHint: op.readOnly === true } },
      async (args: Record<string, unknown>) => {
        try {
          return { content: text(await runShareOp(op, args, remote!)) };
        } catch (error) {
          return { isError: true, content: text({ error: message(error) }) };
        }
      },
    );
  }
  return server;
}

export async function startMcp(): Promise<void> {
  let remote: Remote | undefined;
  try {
    remote = remoteFromEnv(process.env);
  } catch (error) {
    // A URL with no token keeps the local tools and drops the sharing tools.
    console.error(`${message(error)}: the sharing tools are off`);
  }
  await createMcpServer(undefined, remote).connect(new StdioServerTransport());
}
