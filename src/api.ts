import { loadConfig, type Config } from "./config.js";
import { shortId, registerIds } from "./id-map.js";

// --- ANSI colors ---
const c = {
  reset: "\x1b[0m",
  bold: "\x1b[1m",
  dim: "\x1b[2m",
  red: "\x1b[31m",
  green: "\x1b[32m",
  yellow: "\x1b[33m",
  blue: "\x1b[34m",
  magenta: "\x1b[35m",
  cyan: "\x1b[36m",
  white: "\x1b[37m",
  bgRed: "\x1b[41m",
  bgYellow: "\x1b[43m",
};

const CLIENT_INFO =
  "os=mac; osVer=10.15.7; proc=x86; lcid=en-us; deviceType=1; country=us; clientName=skypeteams; clientVer=1415/26011511118";

function headers(config: Config): Record<string, string> {
  return {
    Authentication: `skypetoken=${config.skypeToken}`,
    behavioroverride: "redirectAs404",
    clientinfo: CLIENT_INFO,
    "x-ms-test-user": "False",
  };
}

function baseUrl(config: Config): string {
  if (!config.chatServiceHost) {
    console.error("No chatServiceHost configured. Run: ms-cli auth login");
    process.exit(1);
  }
  return `https://${config.chatServiceHost}/v1`;
}

async function apiGet(path: string, config?: Config): Promise<unknown> {
  const cfg = config ?? loadConfig();
  if (!cfg.skypeToken) {
    console.error("Not logged in. Run: ms-cli auth login");
    process.exit(1);
  }

  const url = `${baseUrl(cfg)}${path}`;
  const res = await fetch(url, { headers: headers(cfg) });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

async function apiPost(path: string, body: unknown, config?: Config): Promise<unknown> {
  const cfg = config ?? loadConfig();
  if (!cfg.skypeToken) {
    console.error("Not logged in. Run: ms-cli auth login");
    process.exit(1);
  }

  const url = `${baseUrl(cfg)}${path}`;
  const res = await fetch(url, {
    method: "POST",
    headers: { ...headers(cfg), "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

async function apiPut(path: string, body: unknown, config?: Config): Promise<unknown> {
  const cfg = config ?? loadConfig();
  if (!cfg.skypeToken) {
    console.error("Not logged in. Run: ms-cli auth login");
    process.exit(1);
  }

  const url = `${baseUrl(cfg)}${path}`;
  const res = await fetch(url, {
    method: "PUT",
    headers: { ...headers(cfg), "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`API error ${res.status}: ${text.slice(0, 300)}`);
  }
  // PUT may return 200 with body or 204 no content
  const ct = res.headers.get("content-type") ?? "";
  if (ct.includes("json")) return res.json();
  return {};
}

// --- Unread helpers ---

/** Parse consumptionhorizon "lastReadId;readTimestamp;???" → lastReadId as bigint */
function parseHorizon(horizon?: string): bigint {
  if (!horizon) return 0n;
  const parts = horizon.split(";");
  try {
    return BigInt(parts[0]);
  } catch {
    return 0n;
  }
}

function isUnread(conv: Conversation): boolean {
  const lastMsgId = conv.lastMessage?.id;
  if (!lastMsgId) return false;
  // Skip system-only messages (ThreadActivity, Event/Call)
  const mtype = conv.lastMessage?.messagetype ?? "";
  if (mtype.startsWith("ThreadActivity/") || mtype === "Event/Call") return false;

  const horizon = parseHorizon(conv.properties?.consumptionhorizon);
  try {
    return BigInt(lastMsgId) > horizon;
  } catch {
    return false;
  }
}

// --- Chat List ---

interface Conversation {
  id: string;
  type: string;
  version: number;
  threadProperties?: {
    topic?: string;
    threadType?: string;
    productThreadType?: string;
  };
  lastMessage?: {
    id?: string;
    imdisplayname?: string;
    content?: string;
    messagetype?: string;
    originalarrivaltime?: string;
  };
  lastUpdatedMessageId?: number;
  properties?: {
    lastimreceivedtime?: string;
    consumptionhorizon?: string;
  };
}

interface ConversationsResponse {
  conversations: Conversation[];
  _metadata: {
    totalCount: number;
    backwardLink?: string;
    syncState?: string;
  };
}

export async function chatList(options: {
  pageSize?: number;
  type?: string;
  unreadOnly?: boolean;
}): Promise<void> {
  const pageSize = options.pageSize ?? 20;
  const data = (await apiGet(
    `/users/ME/conversations?view=mychats&pageSize=${pageSize}`
  )) as ConversationsResponse;

  let conversations = data.conversations;

  if (options.type) {
    const typeFilter = options.type.toLowerCase();
    conversations = conversations.filter((c) => {
      const pt = c.threadProperties?.productThreadType?.toLowerCase() ?? "";
      const tt = c.threadProperties?.threadType?.toLowerCase() ?? "";
      return pt.includes(typeFilter) || tt.includes(typeFilter);
    });
  }

  if (options.unreadOnly) {
    conversations = conversations.filter(isUnread);
  }

  for (const conv of conversations) {
    const topic = conv.threadProperties?.topic ?? "(no topic)";
    const type = conv.threadProperties?.productThreadType ?? conv.threadProperties?.threadType ?? "unknown";
    const lastMsg = conv.lastMessage;
    const lastTime = conv.properties?.lastimreceivedtime ?? lastMsg?.originalarrivaltime ?? "";
    const sender = lastMsg?.imdisplayname ?? "";
    const preview = stripHtml(lastMsg?.content ?? "").slice(0, 60);
    const unread = isUnread(conv);

    const typeColor = type.includes("Channel") ? c.blue : type.includes("Meeting") ? c.magenta : c.cyan;
    const marker = unread ? ` ${c.bgRed}${c.white}${c.bold} UNREAD ${c.reset}` : "";
    const topicStyle = unread ? `${c.bold}${c.white}` : c.dim;

    console.log(`${typeColor}[${type}]${c.reset}${marker} ${topicStyle}${topic}${c.reset}`);
    console.log(`  ${c.dim}id: ${conv.id}${c.reset}`);
    if (sender || preview) {
      console.log(`  ${c.green}${sender}${c.reset}: ${preview} ${c.dim}(${formatTime(lastTime)})${c.reset}`);
    }
    console.log();
  }

  const unreadCount = conversations.filter(isUnread).length;
  const summary = unreadCount > 0
    ? `${c.bold}${conversations.length}${c.reset} conversations (${c.red}${c.bold}${unreadCount} unread${c.reset})`
    : `${conversations.length} conversations (0 unread)`;
  console.log(summary);
}

// --- Chat Read ---

interface Message {
  id: string;
  messagetype: string;
  content?: string;
  imdisplayname?: string;
  from?: string;
  composetime?: string;
  originalarrivaltime?: string;
  rootMessageId?: string;
  properties?: Record<string, unknown>;
}

interface MessagesResponse {
  messages: Message[];
}

export async function chatRead(
  conversationId: string,
  options: { limit?: number; json?: boolean }
): Promise<void> {
  const limit = options.limit ?? 20;
  const data = (await apiGet(
    `/users/ME/conversations/${encodeURIComponent(conversationId)}/messages?view=msnp24Equivalent&pageSize=${limit}`
  )) as MessagesResponse;

  if (options.json) {
    console.log(JSON.stringify(data.messages, null, 2));
    return;
  }

  // Get my consumptionhorizon for this conversation to mark unread messages
  let myHorizon = 0n;
  try {
    const convList = (await apiGet(
      `/users/ME/conversations?view=mychats&pageSize=200`
    )) as ConversationsResponse;
    const conv = convList.conversations.find((c) => c.id === conversationId);
    if (conv) {
      myHorizon = parseHorizon(conv.properties?.consumptionhorizon);
    }
  } catch { }

  // Reverse to show oldest first
  const messages = [...data.messages].reverse();

  // Register message IDs for short ID resolution
  const rootIds = messages.filter((m) => !m.rootMessageId || m.rootMessageId === m.id).map((m) => m.id);
  if (rootIds.length > 0) registerIds(rootIds);

  // Count replies per rootMessageId for thread info
  const threadCounts = new Map<string, number>();
  for (const msg of messages) {
    const root = msg.rootMessageId;
    if (root && root !== msg.id) {
      threadCounts.set(root, (threadCounts.get(root) ?? 0) + 1);
    }
  }

  for (const msg of messages) {
    // Skip system/event messages unless they have useful content
    if (
      msg.messagetype.startsWith("ThreadActivity/") ||
      msg.messagetype === "Event/Call"
    ) {
      const time = formatTime(msg.originalarrivaltime ?? "");
      console.log(`  ${c.dim}[${time}] --- ${msg.messagetype} ---${c.reset}`);
      continue;
    }

    const sender = msg.imdisplayname ?? extractUserId(msg.from) ?? "system";
    const time = formatTime(msg.originalarrivaltime ?? "");
    const content = stripHtml(msg.content ?? "");
    let isNew = false;
    try { isNew = BigInt(msg.id) > myHorizon; } catch { }

    const newTag = isNew ? `${c.red}${c.bold}[NEW]${c.reset} ` : "";

    // Thread info: show reply count for root messages, indent for replies
    const isRoot = !msg.rootMessageId || msg.rootMessageId === msg.id;
    const replyCount = isRoot ? (threadCounts.get(msg.id) ?? 0) : 0;
    const threadTag = replyCount > 0 ? ` ${c.cyan}[${replyCount} replies]${c.reset}` : "";
    const indent = isRoot ? "  " : "    ";
    const replyPrefix = isRoot ? "" : `${c.dim}↳${c.reset} `;
    const msgIdTag = isRoot ? `${c.yellow}${shortId(msg.id)}${c.reset} ` : "";

    console.log(`${indent}${msgIdTag}${c.dim}[${time}]${c.reset} ${newTag}${replyPrefix}${c.green}${c.bold}${sender}${c.reset}: ${content}${threadTag}`);
  }

  const newCount = messages.filter((m) => {
    try { return BigInt(m.id) > myHorizon; } catch { return false; }
  }).length;
  const newSummary = newCount > 0
    ? `(${c.red}${c.bold}${newCount} new${c.reset})`
    : "(0 new)";
  console.log(`\n${messages.length} messages ${newSummary}`);
}

// --- Chat Send ---

export async function chatSend(
  conversationId: string,
  message: string
): Promise<void> {
  const clientMessageId = Date.now().toString() + Math.random().toString(36).slice(2, 8);

  await apiPost(
    `/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`,
    {
      content: message,
      messagetype: "Text",
      contenttype: "text",
      clientmessageid: clientMessageId,
    }
  );

  console.log("Message sent.");
}

// --- Mark as Read ---

export async function chatMarkRead(conversationId: string): Promise<void> {
  // Get latest message ID
  const data = (await apiGet(
    `/users/ME/conversations/${encodeURIComponent(conversationId)}/messages?view=msnp24Equivalent&pageSize=1`
  )) as MessagesResponse;

  if (!data.messages.length) {
    console.log("No messages to mark as read.");
    return;
  }

  const latestId = data.messages[0].id;
  const now = Date.now();

  await apiPut(
    `/users/ME/conversations/${encodeURIComponent(conversationId)}/properties?name=consumptionhorizon`,
    { consumptionhorizon: `${latestId};${now};${latestId}` }
  );

  console.log(`Marked as read up to message ${latestId}.`);
}

// --- Chat Thread ---

export async function chatThread(
  conversationId: string,
  rootMessageId: string,
  options: { limit?: number; json?: boolean }
): Promise<void> {
  // Fetch enough messages to find the thread
  const pageSize = options.limit ?? 200;
  const data = (await apiGet(
    `/users/ME/conversations/${encodeURIComponent(conversationId)}/messages?view=msnp24Equivalent&pageSize=${pageSize}`
  )) as MessagesResponse;

  const thread = data.messages.filter((m) => m.rootMessageId === rootMessageId);

  if (thread.length === 0) {
    console.error("Thread not found. The message may be too old or the ID is incorrect.");
    process.exit(1);
  }

  if (options.json) {
    console.log(JSON.stringify(thread, null, 2));
    return;
  }

  // Show oldest first
  const sorted = [...thread].sort((a, b) => {
    try { return Number(BigInt(a.id) - BigInt(b.id)); } catch { return 0; }
  });

  // First message is the root — show its subject if available
  const root = sorted[0];
  const subject = (root.properties as Record<string, unknown>)?.subject as string | undefined;
  if (subject) {
    console.log(`${c.bold}${c.cyan}── ${subject} ──${c.reset}\n`);
  }

  for (const msg of sorted) {
    if (
      msg.messagetype.startsWith("ThreadActivity/") ||
      msg.messagetype === "Event/Call"
    ) {
      continue;
    }

    const sender = msg.imdisplayname ?? extractUserId(msg.from) ?? "system";
    const time = formatTime(msg.originalarrivaltime ?? "");
    const content = stripHtml(msg.content ?? "");
    const isRoot = msg.id === rootMessageId;
    const indent = isRoot ? "" : "  ";

    console.log(`${indent}${c.dim}[${time}]${c.reset} ${c.green}${c.bold}${sender}${c.reset}: ${content}`);
  }

  console.log(`\n${sorted.length} messages in thread`);
}

// --- Helpers ---

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]*>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function formatTime(iso: string): string {
  if (!iso) return "";
  try {
    const d = new Date(iso);
    const now = new Date();
    const isToday = d.toDateString() === now.toDateString();
    if (isToday) {
      return d.toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" });
    }
    return d.toLocaleDateString("ja-JP", { month: "2-digit", day: "2-digit" }) +
      " " +
      d.toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" });
  } catch {
    return iso;
  }
}

function extractUserId(from?: string): string | undefined {
  if (!from) return undefined;
  const match = from.match(/8:orgid:([a-f0-9-]+)/);
  return match ? match[1] : undefined;
}
