import { join } from "node:path";
import { createFileBlobs, createFileStore } from "./core/file-store.js";
import { PRINCIPAL_KINDS, type PrincipalKind } from "./core/workbook.js";
import type { Ctx } from "./ops.js";

export function ctxFromEnv(env: Record<string, string | undefined>): Ctx {
  const dir = env.SHEET_DIR ?? ".sheets";
  return {
    store: createFileStore(dir),
    blobs: createFileBlobs(join(dir, "blobs")),
    principal() {
      const { SHEET_AS: id = "", SHEET_KIND: kind = "" } = env;
      if (id === "") throw new Error("say who is writing: set --as <id> (or SHEET_AS)");
      if (!PRINCIPAL_KINDS.includes(kind as PrincipalKind)) {
        throw new Error(`say what kind of writer this is: set --kind (or SHEET_KIND) to one of ${PRINCIPAL_KINDS.join(", ")}`);
      }
      return { id, kind: kind as PrincipalKind };
    },
  };
}
