import { spawn, spawnSync, type ChildProcess } from "node:child_process";
import { createServer } from "node:net";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { afterAll, beforeAll, describe, expect, test } from "vitest";
import { createClient } from "./client.js";

// These tests drive the built CLI and MCP server against the real Worker (`wrangler dev`) to share a workbook.
// Failure modes, listed before the code:
// - a sharing command runs with no Worker or no token and does not exit 2, or it prints the token
// - a bad role or kind is sent to the Worker, or a share lands for the wrong kind
// - share does not change what the other principal can read, or unshare leaves the role in place
// - a non-owner can list or change shares, or a refusal exits 0 or is a crash
// - a local MCP server shows sharing tools it cannot serve, or a URL with no token stops the server

const WRANGLER = fileURLToPath(new URL("../node_modules/wrangler/bin/wrangler.js", import.meta.url));
const CLI = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
const KEY_OWNER = "owner-key-0123456789-0123456789-abcdef";
const KEY_VIC = "viewer-key-0123456789-0123456789-abcde";
const KEYS = JSON.stringify({
  [KEY_OWNER]: { id: "ana", kind: "human" },
  [KEY_VIC]: { id: "vic", kind: "human" },
});
const home = mkdtempSync(join(tmpdir(), "wrangler-share-"));

async function freePorts(count: number): Promise<number[]> {
  const open = await Promise.all(
    Array.from({ length: count }, () => new Promise<ReturnType<typeof createServer>>((resolve) => {
      const server = createServer().listen(0, "127.0.0.1", () => resolve(server));
    })),
  );
  const ports = open.map((server) => (server.address() as { port: number }).port);
  await Promise.all(open.map((server) => new Promise((resolve) => server.close(resolve))));
  return ports;
}

let stop = async () => {};
let url = "";
beforeAll(async () => {
  for (let attempt = 1; ; attempt += 1) {
    const [port, inspector] = await freePorts(2);
    const args = [WRANGLER, "dev", "--port", String(port), "--inspector-port", String(inspector), "--ip", "127.0.0.1", "--persist-to", join(home, "state"), "--var", "SHEETS_REGION:unknown", "--var", "SHEETS_UNSAFE_REGION:1", "--var", `SHEETS_API_KEYS:${KEYS}`];
    const env = { ...process.env, CI: "1", NO_COLOR: "1", WRANGLER_SEND_METRICS: "false", XDG_CONFIG_HOME: home, HOME: home };
    const child: ChildProcess = spawn(process.execPath, args, { env, stdio: ["ignore", "pipe", "pipe"], detached: true });
    let [exited, log] = [false, ""];
    child.on("exit", () => (exited = true));
    for (const stream of [child.stdout, child.stderr]) stream?.on("data", (chunk: Buffer) => (log = (log + chunk).slice(-1500)));
    const kill = async () => void (child.pid && !exited && process.kill(-child.pid, "SIGKILL"));
    const base = `http://127.0.0.1:${port}`;
    for (let i = 0; i < 300 && !exited; i += 1) {
      if (await fetch(`${base}/health`).then(() => true, () => false)) {
        [url, stop] = [base, kill];
        return;
      }
      await new Promise((r) => setTimeout(r, 300));
    }
    await kill();
    if (!exited || attempt >= 3) throw new Error(`wrangler dev did not start after ${attempt} tries: ${log}`);
  }
}, 120_000);

afterAll(async () => {
  await stop();
  rmSync(home, { recursive: true, force: true });
});

const env = (token?: string): Record<string, string> => {
  const base: Record<string, string> = { PATH: process.env.PATH ?? "", SHEET_DIR: join(home, "local") };
  return token === undefined ? base : { ...base, SHEET_URL: url, SHEET_TOKEN: token };
};
const cli = (token: string | undefined, ...args: string[]) => {
  const r = spawnSync(process.execPath, [CLI, ...args], { encoding: "utf8", env: env(token) });
  return { code: r.status, out: r.stdout.trim(), err: r.stderr.trim() };
};
const as = (token: string) => createClient({ baseUrl: url, token });
const visible = async (token: string, book: string) => as(token).read({ book, range: "A1" }).then(() => true, () => false);

describe("sharing from the CLI", () => {
  test("it needs a Worker and a token: exit 2, and the token is never printed", () => {
    expect(cli(undefined, "share", "b", "--subject", "vic", "--subject-kind", "human", "--role", "viewer")).toMatchObject({ code: 2 });
    const noToken = spawnSync(process.execPath, [CLI, "shares", "b"], { encoding: "utf8", env: { ...env(), SHEET_URL: url } });
    expect(noToken.status).toBe(2);
    expect(noToken.stderr).toContain("SHEET_TOKEN");
    const bad = cli(KEY_OWNER, "share", "b", "--subject", "vic", "--subject-kind", "robot", "--role", "viewer");
    expect(bad.code).toBe(1);
    expect(bad.err + bad.out).not.toContain(KEY_OWNER);
  });

  test("share, shares, and unshare change what vic can read", async () => {

    await as(KEY_OWNER).create({ book: "plan" });
    expect(await visible(KEY_VIC, "plan")).toBe(false);
    expect(JSON.parse(cli(KEY_OWNER, "share", "plan", "--subject", "vic", "--subject-kind", "human", "--role", "viewer").out)).toEqual({ ok: true });
    expect(await visible(KEY_VIC, "plan")).toBe(true);
    const grants = JSON.parse(cli(KEY_OWNER, "shares", "plan").out).grants as { subjectKind: string; subjectHash: string; role: string }[];
    expect(grants.map((g) => `${g.subjectKind}:${g.role}`).sort()).toEqual(["human:owner", "human:viewer"]);
    expect(JSON.stringify(grants)).not.toContain("vic");
    expect(JSON.parse(cli(KEY_OWNER, "unshare", "plan", "--subject", "vic", "--subject-kind", "human").out)).toEqual({ ok: true });
    expect(await visible(KEY_VIC, "plan")).toBe(false);
    const again = cli(KEY_OWNER, "unshare", "plan", "--subject", "vic", "--subject-kind", "human");
    expect(again.code).toBe(1);
    expect(again.err).toContain("no share");
  });

  test("plain http to a remote host is refused, and shares page", async () => {
    const plain = spawnSync(process.execPath, [CLI, "shares", "b"], { encoding: "utf8", env: { ...env(KEY_OWNER), SHEET_URL: "http://sheets.example.com" } });
    expect(plain.status).toBe(2);
    expect(plain.stderr).toContain("https");
    expect(plain.stderr).not.toContain(KEY_OWNER);
    await as(KEY_OWNER).create({ book: "page" });
    for (let i = 0; i < 5; i += 1) await as(KEY_OWNER).share({ book: "page", kind: "human", subject: `p${i}`, role: "viewer" });
    const seen: string[] = [];
    let cursor: string | undefined;
    for (let pages = 0; pages < 10; pages += 1) {
      const out = JSON.parse(cli(KEY_OWNER, "shares", "page", "--limit", "2", ...(cursor === undefined ? [] : ["--cursor", cursor])).out) as { grants: { subjectHash: string }[]; nextCursor?: string };
      expect(out.grants.length).toBeLessThanOrEqual(2);
      seen.push(...out.grants.map((g) => g.subjectHash));
      cursor = out.nextCursor;
      if (cursor === undefined) break;
    }
    expect(new Set(seen).size).toBe(6); // the owner and five viewers, each once
    expect(cli(KEY_OWNER, "shares", "page", "--limit", "0").code).toBe(1);
  });

  test("invite, invites, link, links, and unlink work from the CLI and never show an email or a token", async () => {
    await as(KEY_OWNER).create({ book: "inv" });
    const invite = JSON.parse(cli(KEY_OWNER, "invite", "inv", "--role", "viewer", "--email", "ada@example.com", "--expires-in-ms", "60000").out) as { id: string };
    expect(invite.id).toBeTruthy();
    expect(cli(KEY_OWNER, "invite", "inv", "--role", "viewer", "--handle", "helper").code).toBe(0);
    const link = JSON.parse(cli(KEY_OWNER, "link", "inv", "--role", "viewer", "--signed-in-only", "true").out) as { id: string; token: string };
    expect(link.token.length).toBeGreaterThan(20);
    const invites = cli(KEY_OWNER, "invites", "inv").out;
    expect(JSON.parse(invites).invites).toHaveLength(2);
    expect(invites).not.toContain("ada@example.com");
    const links = cli(KEY_OWNER, "links", "inv").out;
    expect(JSON.parse(links).links).toMatchObject([{ id: link.id, role: "viewer", signedInOnly: true }]);
    expect(links).not.toContain(link.token);
    expect(links).not.toContain("tokenHash");
    expect(cli(KEY_OWNER, "unlink", "inv", "--link", link.id).out).toBe('{"ok":true}');
    expect(JSON.parse(cli(KEY_OWNER, "links", "inv").out).links[0].revokedAt).toBeTypeOf("number");
    const revokedAt = JSON.parse(cli(KEY_OWNER, "links", "inv").out).links[0].revokedAt;
    expect(cli(KEY_OWNER, "unlink", "inv", "--link", link.id).code).toBe(0);
    expect(JSON.parse(cli(KEY_OWNER, "links", "inv").out).links[0].revokedAt).toBe(revokedAt);
    expect(cli(KEY_OWNER, "link", "inv", "--role", "viewer", "--expires-in-ms").code).toBe(1);
    expect(cli(KEY_OWNER, "link", "inv", "--role", "viewer", "--expires-in-ms", "soon").code).toBe(1);
    expect(cli(KEY_OWNER, "unlink", "inv", "--link", "nope").code).toBe(1);
    expect(cli(KEY_OWNER, "invite", "inv", "--role", "owner", "--handle", "helper").code).toBe(1);
    expect(cli(KEY_OWNER, "invite", "inv", "--role", "viewer").code).toBe(1);
    expect(cli(undefined, "link", "inv", "--role", "viewer").code).toBe(2);
  });

  test("a non-owner cannot invite or make links", () => {
    for (const args of [["invites", "plan"], ["invite", "plan", "--role", "viewer", "--handle", "x"], ["links", "plan"], ["link", "plan", "--role", "viewer"], ["unlink", "plan", "--link", "x"]]) {
      const r = cli(KEY_VIC, ...args);
      expect(r.code, args[0]).toBe(1);
      expect(r.err).toContain("no workbook");
    }
  });

  test("a non-owner cannot list or change shares", () => {
    for (const args of [["shares", "plan"], ["share", "plan", "--subject", "vic", "--subject-kind", "human", "--role", "editor"], ["unshare", "plan", "--subject", "ana", "--subject-kind", "human"]]) {
      const r = cli(KEY_VIC, ...args);
      expect(r.code, args[0]).toBe(1);
      expect(r.err).toContain("no workbook");
    }
  });

  test("--help lists the commands", () => {
    const help = cli(undefined, "--help").out;
    for (const word of ["share <book>", "shares <book>", "unshare <book>", "SHEET_URL", "sheet_share"]) expect(help).toContain(word);
  });
});

describe("sharing from MCP", () => {
  const connect = async (e: Record<string, string>) => {
    const client = new Client({ name: "test", version: "0" });
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [CLI, "mcp"], env: e, stderr: "ignore" }));
    return client;
  };
  const call = async (client: Client, name: string, args: Record<string, unknown>) => {
    const r = (await client.callTool({ name, arguments: args })) as { isError?: boolean; content: { text: string }[] };
    return { isError: r.isError === true, body: JSON.parse(r.content[0].text) };
  };

  test("a local server, and a URL with no token, have no sharing tools", async () => {
    for (const e of [env(), { ...env(), SHEET_URL: url }]) {
      const client = await connect(e);
      const names = (await client.listTools()).tools.map((t) => t.name);
      await client.close();
      expect(names).toContain("sheet_new");
      expect(names).not.toContain("sheet_share");
    }
  });

  test("with a Worker it has sheet_share, sheet_shares, and sheet_unshare, and refusals are tool errors", async () => {
    await as(KEY_OWNER).create({ book: "mcp" });
    const client = await connect(env(KEY_OWNER));
    try {
      const names = (await client.listTools()).tools.map((t) => t.name);
      expect(names).toEqual(expect.arrayContaining(["sheet_share", "sheet_shares", "sheet_unshare"]));
      expect((await call(client, "sheet_share", { book: "mcp", subject: "bot", subjectKind: "agent", role: "editor" })).body).toEqual({ ok: true });
      expect((await call(client, "sheet_shares", { book: "mcp" })).body.grants.map((g: { subjectKind: string; role: string }) => `${g.subjectKind}:${g.role}`).sort()).toEqual(["agent:editor", "human:owner"]);
      expect((await call(client, "sheet_unshare", { book: "mcp", subject: "bot", subjectKind: "agent" })).body).toEqual({ ok: true });
      expect((await call(client, "sheet_unshare", { book: "mcp", subject: "bot", subjectKind: "agent" })).isError).toBe(true);
      expect((await call(client, "sheet_shares", { book: "ghost" })).isError).toBe(true);
      expect(names).toEqual(expect.arrayContaining(["sheet_invite", "sheet_invites", "sheet_link", "sheet_links", "sheet_unlink"]));
      expect((await call(client, "sheet_invite", { book: "mcp", role: "viewer", handle: "helper" })).isError).toBe(false);
      const made = (await call(client, "sheet_link", { book: "mcp", role: "viewer", signedInOnly: true })).body as { id: string; token: string };
      const listed = await call(client, "sheet_links", { book: "mcp" });
      expect(listed.body.links).toMatchObject([{ id: made.id, signedInOnly: true }]);
      expect(JSON.stringify(listed.body)).not.toContain(made.token);
      expect((await call(client, "sheet_invites", { book: "mcp" })).body.invites).toMatchObject([{ for: "handle", handle: "helper" }]);
      expect((await call(client, "sheet_unlink", { book: "mcp", link: made.id })).body).toEqual({ ok: true });
      const owner = (await client.callTool({ name: "sheet_link", arguments: { book: "mcp", role: "owner" } })) as { isError?: boolean };
      expect(owner.isError).toBe(true);
    } finally {
      await client.close();
    }
  });
});
