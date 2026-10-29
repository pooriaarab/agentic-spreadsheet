import { z } from "zod";
import { createClient, type SheetsClient } from "./client.js";
import { SHARE_ROLES, SUBJECT_KIND_NAMES } from "./access.js";

// Roles live on the Worker, so the sharing commands and tools call its HTTP API. A local store has no roles.
export interface Remote {
  url: string;
  token: string;
}

// SHEET_URL names the Worker and SHEET_TOKEN is the credential. A URL with no token is an error.
export function remoteFromEnv(env: Record<string, string | undefined>): Remote | undefined {
  const { SHEET_URL: url = "", SHEET_TOKEN: token = "" } = env;
  if (url === "") return undefined;
  if (token === "") throw new Error("sharing needs SHEET_TOKEN, the credential for the Worker at SHEET_URL");
  return { url, token };
}

const book = z.string().min(1).describe("Workbook id.");
const subject = z.string().min(1).describe("The id of the person or agent.");
const subjectKind = z.enum(SUBJECT_KIND_NAMES).describe("A person (human) and an agent with one id hold separate roles.");

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
    name: "share", description: "Give a person or an agent a role on a workbook, through the Worker. Owner only.", args: ["book"],
    shape: { book, subject, subjectKind, role: z.enum(SHARE_ROLES).describe("viewer, commenter, or editor.") },
    run: (a: { book: string; subject: string; subjectKind: "human" | "agent"; role: "viewer" | "commenter" | "editor" }, client) =>
      client.share({ book: a.book, kind: a.subjectKind, subject: a.subject, role: a.role }),
  },
  {
    name: "shares", description: "List the roles on a workbook, through the Worker. Owner only. Subject ids are hashes.", args: ["book"],
    shape: { book }, readOnly: true,
    run: (a: { book: string }, client) => client.shares({ book: a.book }),
  },
  {
    name: "unshare", description: "Take a role away on a workbook, through the Worker. Owner only.", args: ["book"],
    shape: { book, subject, subjectKind },
    run: (a: { book: string; subject: string; subjectKind: "human" | "agent" }, client) =>
      client.unshare({ book: a.book, kind: a.subjectKind, subject: a.subject }),
  },
];

export const runShareOp = (op: ShareOp, args: unknown, remote: Remote) => op.run(args, createClient({ baseUrl: remote.url, token: remote.token }));
