import { useEffect, useRef, useState } from "react";
import { openDocument, peerLabel } from "agentic-realtime/browser";
import type { Peer as RoomPeer } from "agentic-realtime/client";
import type { PresenceEntry } from "agentic-design/react";
import { cursorOf } from "./grid.js";
import type { RoomPresence } from "./host.js";

export interface Peer {
  entry: PresenceEntry;
  cursor: string | null;
  label: string;
}
export type Live = "connecting" | "connected" | "disconnected";

const base64url = (text: string) => btoa(String.fromCharCode(...new TextEncoder().encode(text))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");

// A browser WebSocket cannot send an Authorization header, and a URL reaches logs. The key goes in a subprotocol entry instead.
class TokenSocket extends WebSocket {
  constructor(url: string | URL, options?: { headers?: Record<string, string> }) {
    const token = /^Bearer (.+)$/.exec(options?.headers?.authorization ?? "")?.[1];
    super(url, token ? ["agentic-spreadsheet", `agentic-spreadsheet.bearer.${base64url(token)}`] : undefined);
  }
}

const toPeer = (peer: RoomPeer): Peer => ({
  entry: { id: String(peer.clientId), name: peer.name, kind: peer.machine ? "agent" : "person", status: "active" },
  cursor: cursorOf(peer.cursor),
  label: peerLabel(peer),
});

// Join the presence room of a workbook. The room carries who is here and where their cursor is.
// Cells do not travel through it: the API stores them. With `host`, the app's own room is used, and no room is opened.
export function usePresence(book: string, token: string | undefined, cursor: string, name = "Me", host?: RoomPresence) {
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
    const url = `${scheme}://${location.host}/api/v1/docs/${encodeURIComponent(book)}/ws`;
    openDocument({ url, headers: { authorization: `Bearer ${token}` }, WebSocket: TokenSocket, storageName: `presence:${book}`, presence: { name, cursor: latest.current }, onStatus: (status) => !closed && setLive(status) })
      .then(({ client: opened }) => {
        if (closed) return void opened.disconnect();
        client.current = opened;
        opened.setPresence({ cursor: latest.current });
        stop = opened.onPeers((list) => setPeers(list.map(toPeer)));
      })
      .catch(() => !closed && setLive("disconnected"));
    return () => {
      closed = true;
      stop();
      void client.current?.disconnect();
      client.current = null;
    };
  }, [book, token, name]);

  // The app's room. The cursor leaves it with the sheet, and the room itself stays open.
  useEffect(() => {
    if (!host) return;
    setPeers(host.peers.map(toPeer));
    const stop = host.onPeers((list) => setPeers(list.map(toPeer)));
    return () => {
      stop();
      host.setPresence({ cursor: null });
    };
  }, [host]);

  useEffect(() => (host ?? client.current)?.setPresence({ cursor }), [cursor, host]);
  return { peers, live };
}
