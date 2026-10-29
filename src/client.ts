// A typed client for the Worker HTTP API. It uses only `fetch`, so it runs in browsers, Workers, and Node 22.
// It never retries: a write that fails may or may not have landed, and the caller decides what to do.

export type Cell = string | number | boolean | null;
/** A formula that fails reads as this object, such as `{ error: "#REF!" }`, never as a string. */
export type FormulaError = { error: "#REF!" | "#DIV/0!" | "#NAME?" | "#VALUE!" | "#NUM!" | "#N/A" | "#CIRCULAR" | "#DEPTH!" };
/** A computed value: what a cell holds, or the error of its formula. */
export type Value = Cell | FormulaError;
export type PrincipalKind = "human" | "agent_of_human" | "agent_of_agent" | "agent" | "user_token";
export type SubjectKind = "human" | "agent";
export type ShareRole = "viewer" | "commenter" | "editor";

export interface CellAuthor {
  id: string;
  kind: PrincipalKind;
  machine: boolean;
  at: string;
}

export interface ClientOptions {
  /** The origin of the Worker, such as `https://sheets.example.com`. A path prefix is kept. */
  baseUrl: string;
  /** The API key. It goes in an `authorization: Bearer` header. */
  token: string;
  /** Use this in place of the global `fetch`. */
  fetch?: typeof fetch;
}

/** Every call takes an optional signal. An aborted call rejects with the error that `fetch` throws. */
export interface CallOptions {
  signal?: AbortSignal;
}

export interface SheetTarget {
  /** A sheet name. The first sheet is the default. */
  sheet?: string;
}

export interface ReadResult {
  sheet: string;
  range: string;
  values: Value[][];
  formulas: (string | null)[][];
  authors: (CellAuthor | null)[][];
}

export type ErrorCode = "bad_request" | "unauthorized" | "forbidden" | "not_found" | "conflict" | "too_large" | "server_error" | "unknown";

const CODES: Record<number, ErrorCode> = { 400: "bad_request", 401: "unauthorized", 403: "forbidden", 404: "not_found", 409: "conflict", 413: "too_large" };

/** The API answers with a status and `{ error }`. `code` names the status, so a caller can switch on it. */
export class SheetsApiError extends Error {
  readonly status: number;
  readonly code: ErrorCode;
  constructor(status: number, message: string) {
    super(message);
    this.name = "SheetsApiError";
    this.status = status;
    this.code = CODES[status] ?? (status >= 500 ? "server_error" : "unknown");
  }
}

const segment = encodeURIComponent;

export function createClient({ baseUrl, token, fetch: fetchImpl }: ClientOptions) {
  if (typeof token !== "string" || token === "") throw new TypeError("createClient needs a token");
  const origin = baseUrl.replace(/\/+$/, "");
  const root = `${origin}/api/v1/sheets`;

  // The message comes from `{ error }` when the body has one.
  async function failed(res: Response): Promise<SheetsApiError> {
    let message = `the API answered ${res.status}`;
    try {
      const body: unknown = await res.json();
      if (typeof body === "object" && body !== null && typeof (body as { error?: unknown }).error === "string") message = (body as { error: string }).error;
    } catch {
      // The body is not JSON. Keep the status message.
    }
    return new SheetsApiError(res.status, message);
  }

  async function send(method: string, path: string, opts: CallOptions, init: { query?: Record<string, string | number | undefined>; body?: unknown } = {}): Promise<Response> {
    const query = new URLSearchParams();
    for (const [key, value] of Object.entries(init.query ?? {})) if (value !== undefined) query.set(key, String(value));
    const text = query.toString();
    const qs = text === "" ? "" : `?${text}`;
    const headers: Record<string, string> = { authorization: `Bearer ${token}` };
    if (init.body !== undefined) headers["content-type"] = "application/json";
    // Looked up on each call, so a test that replaces the global `fetch` is seen.
    const res = await (fetchImpl ?? fetch)(`${root}${path}${qs}`, { method, headers, body: init.body === undefined ? undefined : JSON.stringify(init.body), signal: opts.signal });
    if (res.ok) return res;
    throw await failed(res);
  }
  async function json<T>(method: string, path: string, opts: CallOptions, init?: Parameters<typeof send>[3]): Promise<T> {
    return (await send(method, path, opts, init)).json() as Promise<T>;
  }
  const book = (id: string) => `/${segment(id)}`;

  return {
    /** `GET /`. Pass `limit` for a page, and the `cursor` of the last page for the next. */
    list: (args: { limit?: number; cursor?: string } = {}, opts: CallOptions = {}) =>
      json<{ books: string[]; cursor?: string }>("GET", "", opts, { query: args }),
    /** `POST /`. */
    create: (args: { book: string; sheet?: string }, opts: CallOptions = {}) => json<{ book: string; sheets: string[] }>("POST", "", opts, { body: args }),
    /** `GET /:book/range`. */
    read: ({ book: id, ...query }: { book: string; range: string } & SheetTarget, opts: CallOptions = {}) => json<ReadResult>("GET", `${book(id)}/range`, opts, { query }),
    /** `PUT /:book/range`. A string that starts with `=` is a formula. */
    write: ({ book: id, ...body }: { book: string; anchor: string; values: Cell[][] } & SheetTarget, opts: CallOptions = {}) =>
      json<{ ok: true; rows: number }>("PUT", `${book(id)}/range`, opts, { body }),
    /** `PUT /:book/formula`. */
    setFormula: ({ book: id, ...body }: { book: string; cell: string; formula: string } & SheetTarget, opts: CallOptions = {}) =>
      json<{ cell: string; formula: string; value: Value }>("PUT", `${book(id)}/formula`, opts, { body }),
    /** `POST /:book/rows`. */
    insertRows: ({ book: id, ...body }: { book: string; at: number; count?: number } & SheetTarget, opts: CallOptions = {}) =>
      json<{ ok: true }>("POST", `${book(id)}/rows`, opts, { body }),
    /** `DELETE /:book/rows`. A formula that points at a deleted cell reads `#REF!`. */
    deleteRows: ({ book: id, ...query }: { book: string; at: number; count?: number } & SheetTarget, opts: CallOptions = {}) =>
      json<{ ok: true }>("DELETE", `${book(id)}/rows`, opts, { query }),
    /** `POST /:book/cols`. `at` is a column number, and A is 1. */
    insertCols: ({ book: id, ...body }: { book: string; at: number; count?: number } & SheetTarget, opts: CallOptions = {}) =>
      json<{ ok: true }>("POST", `${book(id)}/cols`, opts, { body }),
    /** `DELETE /:book/cols`. */
    deleteCols: ({ book: id, ...query }: { book: string; at: number; count?: number } & SheetTarget, opts: CallOptions = {}) =>
      json<{ ok: true }>("DELETE", `${book(id)}/cols`, opts, { query }),
    /** `GET /:book/sheets`. Each sheet comes with its used range, or `null` when it is empty. */
    sheets: ({ book: id }: { book: string }, opts: CallOptions = {}) => json<{ sheets: { name: string; extent: string | null }[] }>("GET", `${book(id)}/sheets`, opts),
    /** `POST /:book/sheets`. */
    addSheet: ({ book: id, ...body }: { book: string; name: string }, opts: CallOptions = {}) => json<{ ok: true; sheets: string[] }>("POST", `${book(id)}/sheets`, opts, { body }),
    /** `DELETE /:book/sheets/:name`. The last sheet stays. */
    deleteSheet: ({ book: id, name }: { book: string; name: string }, opts: CallOptions = {}) =>
      json<{ ok: true; sheets: string[] }>("DELETE", `${book(id)}/sheets/${segment(name)}`, opts),
    /** `GET /:book/export`. Resolves to the CSV text. */
    exportCsv: async ({ book: id, ...query }: { book: string } & SheetTarget, opts: CallOptions = {}) => (await send("GET", `${book(id)}/export`, opts, { query })).text(),
    /** `POST /:book/import`. */
    importCsv: ({ book: id, ...body }: { book: string; csv: string; asText?: boolean } & SheetTarget, opts: CallOptions = {}) =>
      json<{ ok: true; rows: number }>("POST", `${book(id)}/import`, opts, { body }),
    /** `PUT /:book/share`. Only the owner may share. */
    share: ({ book: id, ...body }: { book: string; kind: SubjectKind; subject: string; role: ShareRole }, opts: CallOptions = {}) =>
      json<{ ok: true }>("PUT", `${book(id)}/share`, opts, { body }),
    /** `GET /:book/shares`. Only the owner may list. A subject shows as a hash of its id. Up to 100 grants come back: pass `nextCursor` as `cursor` to read on. */
    shares: ({ book: id, ...query }: { book: string; limit?: number; cursor?: string }, opts: CallOptions = {}) =>
      json<{ grants: { subjectKind: SubjectKind; subjectHash: string; role: "owner" | ShareRole }[]; nextCursor?: string }>("GET", `${book(id)}/shares`, opts, { query }),
    /** `DELETE /:book/share/:kind/:subject`. */
    unshare: (args: { book: string; kind: SubjectKind; subject: string }, opts: CallOptions = {}) =>
      json<{ ok: true }>("DELETE", `${book(args.book)}/share/${segment(args.kind)}/${segment(args.subject)}`, opts),
    /** `DELETE /:book`. Only the owner may delete. */
    remove: (args: { book: string }, opts: CallOptions = {}) => json<{ ok: true }>("DELETE", book(args.book), opts),
    /** `GET /health`. It needs no key and sits outside `/api/v1/sheets`. */
    health: async (opts: CallOptions = {}) => {
      const res = await (fetchImpl ?? fetch)(`${origin}/health`, { signal: opts.signal });
      if (!res.ok) throw await failed(res);
      return res.json() as Promise<{ ok: true; pack: string; version: string }>;
    },
  };
}

export type SheetsClient = ReturnType<typeof createClient>;
