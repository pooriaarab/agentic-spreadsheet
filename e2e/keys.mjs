export const ANA = "human-key-ana-0123456789-0123456789-abcdef";
export const BEN = "human-key-ben-0123456789-0123456789-abcdef";
export const BOT = "agent-key-bot-0123456789-0123456789-abcde";
export const KEYS = {
  [ANA]: { id: "ana", kind: "human" },
  [BEN]: { id: "ben", kind: "human" },
  [BOT]: { id: "bot-1", kind: "agent_of_human", parent: "ana" },
};
