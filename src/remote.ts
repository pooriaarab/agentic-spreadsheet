import { z } from "zod";
import { createClient, type SheetsClient } from "./client.js";
import { SHARE_ROLES, SUBJECT_KIND_NAMES } from "./access.js";

// Roles live on the Worker, so the sharing commands and tools call its HTTP API. A local store has no roles.
export interface Remote {
  url: string;
  token: string;
}

// A fault in how sharing is set up, not a refusal from the Worker. The CLI exits 2 for it.
export class SetupError extends Error {}

// SHEET_URL names the Worker and SHEET_TOKEN is the credential. A URL with no token is an error.
export function remoteFromEnv(env: Record<string, string | undefined>): Remote | undefined {
  const { SHEET_URL: url = "", SHEET_TOKEN: token = "" } = env;
  if (url === "") return undefined;
  if (token === "") throw new SetupError("sharing needs SHEET_TOKEN, the credential for the Worker at SHEET_URL");
  // The key goes in a header, so only a loopback host may use plain http.
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    throw new SetupError("SHEET_URL must be an absolute URL");
  }
  const loopback = ["localhost", "127.0.0.1", "[::1]"].includes(parsed.hostname);
  if (parsed.protocol !== "https:" && !(parsed.protocol === "http:" && loopback)) throw new SetupError("SHEET_URL must use https, except for a loopback host");
  return { url, token };
}

const book = z.string().min(1).describe("Workbook id.");
const cursor = z.string().optional().describe("The nextCursor of the last page.");
const subject = z.string().min(1).describe("The id of the person or agent.");
const subjectKind = z.enum(SUBJECT_KIND_NAMES).describe("A person (human) and an agent with one id hold separate roles.");

// The CLI gives text, and a flag with no value gives `true`: only a number or digits pass.
const whole = (max: number) => z.union([z.number(), z.string().regex(/^\d{1,15}$/).transform(Number)]).pipe(z.number().int().min(1).max(max)).optional();
const lifetime = whole(Number.MAX_SAFE_INTEGER).describe("How long it lives, in milliseconds.");
// The CLI gives "true" or "false" as text, and MCP gives a boolean.
const flag = z.union([z.boolean(), z.enum(["true", "false"]).transform((v) => v === "true")]).optional().describe("Only signed-in people may use it.");

export interface ShareOp {
  name: string;
  description: string;
  args: string[];
  shape: z.ZodRawShape;
  readOnly?: boolean;
  run(args: any, client: SheetsClient): Promise<unknown>;
}

export const SHARE_OPS: ShareOp[] = [
  {
    name: "share", description: "Give a person or an agent a role on a workbook, through the Worker. Owner only. The workbook must exist on the Worker: the other tools use the local store.", args: ["book"],
    shape: { book, subject, subjectKind, role: z.enum(SHARE_ROLES).describe("viewer, commenter, or editor.") },
    run: (a: { book: string; subject: string; subjectKind: "human" | "agent"; role: "viewer" | "commenter" | "editor" }, client) =>
      client.share({ book: a.book, kind: a.subjectKind, subject: a.subject, role: a.role }),
  },
  {
    name: "shares", description: "List the roles on a workbook, through the Worker. Owner only. Subject ids are hashes. Up to 100 a page: pass nextCursor back as cursor.", args: ["book"],
    shape: { book, limit: z.coerce.number().int().min(1).max(100).optional().describe("Roles a page, 100 at most."), cursor }, readOnly: true,
    run: (a: { book: string; limit?: number; cursor?: string }, client) => client.shares(a),
  },
  {
    name: "unshare", description: "Take a role away on a workbook, through the Worker. Owner only.", args: ["book"],
    shape: { book, subject, subjectKind },
    run: (a: { book: string; subject: string; subjectKind: "human" | "agent" }, client) =>
      client.unshare({ book: a.book, kind: a.subjectKind, subject: a.subject }),
  },
  {
    name: "invites", description: "List the invites on a workbook, through the Worker. Owner only. An invite shows no email. Up to 100 a page: pass nextCursor back as cursor.", args: ["book"],
    shape: { book, limit: whole(100).describe("Invites a page, 100 at most."), cursor }, readOnly: true,
    run: (a: { book: string; limit?: number; cursor?: string }, client) => client.invites(a),
  },
  {
    name: "invite", description: "Invite a person by email or an agent by handle to a workbook, through the Worker. Give email or handle, not both. The role is not owner. Owner only.", args: ["book"],
    shape: { book, role: z.enum(SHARE_ROLES).describe("viewer, commenter, or editor."), email: z.string().min(3).optional(), handle: z.string().min(1).optional(), expiresInMs: lifetime },
    run: (a: { book: string; role: "viewer" | "commenter" | "editor"; email?: string; handle?: string; expiresInMs?: number }, client) => client.invite(a),
  },
  {
    name: "links", description: "List the share links on a workbook, through the Worker. Owner only. A link shows no token. Up to 100 a page: pass nextCursor back as cursor.", args: ["book"],
    shape: { book, limit: whole(100).describe("Links a page, 100 at most."), cursor }, readOnly: true,
    run: (a: { book: string; limit?: number; cursor?: string }, client) => client.links(a),
  },
  {
    name: "link", description: "Make a share link for a workbook, through the Worker. The reply holds the token once: it cannot be read again. The role is not owner. Owner only.", args: ["book"],
    shape: { book, role: z.enum(SHARE_ROLES).describe("viewer, commenter, or editor."), signedInOnly: flag, expiresInMs: lifetime },
    run: (a: { book: string; role: "viewer" | "commenter" | "editor"; signedInOnly?: boolean; expiresInMs?: number }, client) => client.link(a),
  },
  {
    name: "unlink", description: "Revoke a share link of a workbook, through the Worker. Owner only.", args: ["book"],
    shape: { book, link: z.string().min(1).describe("Link id.") },
    run: (a: { book: string; link: string }, client) => client.unlink(a),
  },
];

export const runShareOp = (op: ShareOp, args: unknown, remote: Remote) => op.run(args, createClient({ baseUrl: remote.url, token: remote.token }));
