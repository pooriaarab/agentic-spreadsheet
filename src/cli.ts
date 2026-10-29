#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { z } from "zod";
import { ctxFromEnv } from "./env.js";
import { message, OPS } from "./ops.js";
import { remoteFromEnv, runShareOp, SetupError, SHARE_OPS } from "./remote.js";
import { VERSION } from "./version.js";

const HELP = `agentic-spreadsheet ${VERSION}
Usage: agentic-spreadsheet <command> [args] [flags]

Commands:
${OPS.map((op) => `  ${`${op.name} ${op.args.map((a) => `<${a}>`).join(" ")}`.trim().padEnd(26)}${op.description}`).join("\n")}
${SHARE_OPS.map((op) => `  ${`${op.name} ${op.args.map((a) => `<${a}>`).join(" ")}`.trim().padEnd(26)}${op.description}`).join("\n")}
  mcp                       Start the MCP server on stdio.

Flags: --sheet <name> --dir <path> --file <csv> --as-text.
A write needs --as <id> and --kind <kind>, or SHEET_AS and SHEET_KIND. The kind is one of
human, agent_of_human, agent_of_agent, agent, user_token. There is no default.
share, shares, unshare, invite, invites, link, links, and unlink call a Worker: set SHEET_URL and SHEET_TOKEN. share and unshare take
--subject <id> --subject-kind <human|agent> (--kind is the kind of a writer), and share takes --role <viewer|commenter|editor>.
invite takes --role and --email <e> or --handle <h>. link takes --role, --signed-in-only <true|false>. Both take --expires-in-ms <n>.
unlink takes --link <id>. A local store has no roles. With SHEET_URL and SHEET_TOKEN set, "mcp" also has sheet_share, sheet_shares, sheet_unshare, sheet_invite, sheet_invites, sheet_link, sheet_links, and sheet_unlink.
Env: SHEET_URL, SHEET_TOKEN, SHEET_DIR, SHEET_AS, SHEET_KIND. The store is ./.sheets unless --dir or SHEET_DIR says otherwise.`;

async function main(argv: string[]): Promise<number> {
  const [command, ...rest] = argv;
  if (command === undefined || command === "--help" || command === "-h") return (console.log(HELP), 0);
  if (command === "--version" || command === "-V") return (console.log(VERSION), 0);
  if (command === "mcp") {
    if (rest.includes("--help") || rest.includes("-h")) return (console.log(HELP), 0);
    if (rest.includes("--version") || rest.includes("-V")) return (console.log(VERSION), 0);
    await (await import("./mcp.js")).startMcp();
    return 0;
  }
  const shareOp = SHARE_OPS.find((o) => o.name === command);
  const op = shareOp ?? OPS.find((o) => o.name === command);
  if (op === undefined) return (console.error(`unknown command "${command}". Run with --help.`), 2);
  const flags: Record<string, string | boolean> = {};
  const positional: string[] = [];
  for (let i = 0; i < rest.length; i += 1) {
    if (!rest[i].startsWith("--")) positional.push(rest[i]);
    else if (rest[i + 1] === undefined || rest[i + 1].startsWith("--")) flags[rest[i].slice(2)] = true;
    else flags[rest[i].slice(2)] = rest[(i += 1)];
  }
  try {
    const { dir, as, kind, file, ...named } = flags;
    const args: Record<string, string | boolean> = Object.fromEntries(
      Object.entries(named).map(([key, value]) => [key.replace(/-(\w)/g, (_, c: string) => c.toUpperCase()), value]),
    );
    const env = { ...process.env, ...(dir && { SHEET_DIR: String(dir) }), ...(as && { SHEET_AS: String(as) }), ...(kind && { SHEET_KIND: String(kind) }) };
    op.args.forEach((name, i) => positional[i] !== undefined && (args[name] = positional[i]));
    if (typeof file === "string") args.csv = readFileSync(file, "utf8");
    if (shareOp !== undefined) {
      const remote = remoteFromEnv(process.env);
      if (remote === undefined) return (console.error("sharing needs a Worker: set SHEET_URL and SHEET_TOKEN. A local store has no roles."), 2);
      console.log(JSON.stringify(await runShareOp(shareOp, z.object(shareOp.shape).strict().parse(args), remote)));
      return 0;
    }
    const result = (await (op as (typeof OPS)[number]).run(z.object(op.shape).strict().parse(args), ctxFromEnv(env))) as { csv?: string };
    if (op.name === "export") process.stdout.write(result.csv ?? "");
    else console.log(JSON.stringify(result));
    return 0;
  } catch (error) {
    console.error(`error: ${message(error)}`);
    return error instanceof SetupError ? 2 : 1;
  }
}

main(process.argv.slice(2)).then(
  (code) => void (code !== 0 && process.exit(code)),
  (error: unknown) => (console.error(`error: ${message(error)}`), process.exit(1)),
);
