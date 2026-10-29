// The Worker that `wrangler dev` runs: the API, the web UI files, and the presence room.
// It lives apart from `worker.ts` because the room code needs the Workers runtime, and the tests import `worker.ts` in Node.
import { createSyncWorker, SyncRoom } from "agentic-realtime/worker";
import { acceptWebSocketProtocol } from "agentic-identity/http";
import api, { roomGrant, WS_PROTOCOL, type WorkerEnv } from "./worker.js";

export { SyncRoom };

type Env = WorkerEnv & { SYNC_ROOM: unknown; ASSETS: { fetch(request: Request): Promise<Response> } };

const sync = createSyncWorker({
  authorize: (request, book, env) => roomGrant(request, book, env as WorkerEnv),
});
// Only the socket route is open. The room also has admin routes, and they stay shut.
const SOCKET = /^\/api\/v1\/docs\/[^/]+\/ws$/;

export default {
  async fetch(request: Request, env: Env, ctx: unknown): Promise<Response> {
    const { pathname } = new URL(request.url);
    if (SOCKET.test(pathname)) {
      const room = (await sync.fetch!(request as never, env as never, ctx as never)) as unknown as Response;
      return acceptWebSocketProtocol(request, room, WS_PROTOCOL);
    }
    if (pathname.startsWith("/api/") || pathname === "/health") return api.fetch(request, env);
    return env.ASSETS.fetch(request);
  },
};
