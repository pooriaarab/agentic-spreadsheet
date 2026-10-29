import { useEffect, useRef, useState } from "react";
import { openDocument, peerLabel } from "agentic-realtime/browser";
import type { PresenceEntry } from "agentic-design/react";
import { cursorOf } from "./grid.js";

export interface Peer {
  entry: PresenceEntry;
  cursor: string | null;
  label: string;
}
export type Live = "connecting" | "connected" | "disconnected";

// Join the presence room of a workbook. The room carries who is here and where their cursor is.
// Cells do not travel through it: the API stores them.
export function usePresence(book: string, token: string | undefined, cursor: string, name = "Me") {
  const [peers, setPeers] = useState<Peer[]>([]);
  const [live, setLive] = useState<Live>(token === undefined ? "disconnected" : "connecting");
  const client = useRef<Awaited<ReturnType<typeof openDocument>>["client"] | null>(null);
  const latest = useRef(cursor);
  latest.current = cursor;

  useEffect(() => {
    if (token === undefined) return;
    let closed = false;
    let stop = () => {};
    setLive("connecting");
    setPeers([]);
    const scheme = location.protocol === "https:" ? "wss" : "ws";
    const url = `${scheme}://${location.host}/api/v1/docs/${encodeURIComponent(book)}/ws?token=${encodeURIComponent(token)}`;
    openDocument({ url, storageName: `presence:${book}`, presence: { name, cursor: latest.current }, onStatus: (status) => !closed && setLive(status) })
      .then(({ client: opened }) => {
        if (closed) return void opened.disconnect();
        client.current = opened;
        opened.setPresence({ cursor: latest.current });
        stop = opened.onPeers((list) =>
          setPeers(
            list.map((peer) => ({
              entry: { id: String(peer.clientId), name: peer.name, kind: peer.machine ? "agent" : "person", status: "active" },
              cursor: cursorOf(peer.cursor),
              label: peerLabel(peer),
            })),
          ),
        );
      })
      .catch(() => !closed && setLive("disconnected"));
    return () => {
      closed = true;
      stop();
      void client.current?.disconnect();
      client.current = null;
    };
  }, [book, token, name]);

  useEffect(() => client.current?.setPresence({ cursor }), [cursor]);
  return { peers, live };
}
