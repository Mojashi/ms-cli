import { loadConfig, getCurrentAccountKey, type Config } from "./config.js";
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

/** Print which account/tenant the following output belongs to. */
export function printAccountHeader(): void {
  const cfg = loadConfig();
  const key = getCurrentAccountKey();
  const org = cfg.tenantName ?? key ?? "unknown";
  const parts = [org];
  if (key && key !== org) parts.push(key);
  if (cfg.tenantId) parts.push(`tenant ${cfg.tenantId.slice(0, 8)}`);
  if (cfg.region) parts.push(cfg.region);
  console.log(`${c.dim}${c.bold}── ${parts.join(" │ ")} ──${c.reset}`);
}

export interface ChatEntry {
  sortTs: number;
  unread: boolean;
  lines: string[];
}

/** Fetch + format conversations into sortable entries (no header/summary printed). */
export async function chatListEntries(options: {
  pageSize?: number;
  type?: string;
  unreadOnly?: boolean;
  tag?: string;
}): Promise<ChatEntry[]> {
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

  const tag = options.tag ? `${c.dim}[${options.tag}]${c.reset} ` : "";

  return conversations.map((conv) => {
    const topic = conv.threadProperties?.topic ?? "(no topic)";
    const type = conv.threadProperties?.productThreadType ?? conv.threadProperties?.threadType ?? "unknown";
    const lastMsg = conv.lastMessage;
    const lastTime = conv.properties?.lastimreceivedtime ?? lastMsg?.originalarrivaltime ?? "";
    const sender = lastMsg?.imdisplayname ?? "";
    const rawContent = lastMsg?.content ?? "";
    const msgtype = lastMsg?.messagetype ?? "";
    const unread = isUnread(conv);

    const isReaction = msgtype === "MessageReaction" || msgtype === "Signal/Flamingo";
    let preview: string;
    if (isReaction) {
      const keyMatch = rawContent.match(/key="([^"]+)"/);
      const emoji = keyMatch ? (reactionEmoji[keyMatch[1]] ?? keyMatch[1]) : "\u{1F44D}";
      const textPreview = stripHtml(rawContent).slice(0, 40);
      preview = `${c.dim}[reaction: ${emoji}]${c.reset}${textPreview ? ` on: "${textPreview}"` : ""}`;
    } else {
      preview = stripHtml(rawContent).slice(0, 60);
    }

    const typeColor = type.includes("Channel") ? c.blue : type.includes("Meeting") ? c.magenta : c.cyan;
    const marker = unread ? ` ${c.bgRed}${c.white}${c.bold} UNREAD ${c.reset}` : "";
    const topicStyle = unread ? `${c.bold}${c.white}` : c.dim;

    const lines = [
      `${tag}${typeColor}[${type}]${c.reset}${marker} ${topicStyle}${topic}${c.reset}`,
      `  ${c.dim}id: ${conv.id}${c.reset}`,
    ];
    if (sender || preview) {
      lines.push(`  ${c.green}${sender}${c.reset}: ${preview} ${c.dim}(${formatTime(lastTime)})${c.reset}`);
    }

    const ts = Date.parse(lastTime);
    return { sortTs: Number.isNaN(ts) ? 0 : ts, unread, lines };
  });
}

/** Print a list of chat entries (newest first), with a count summary. */
export function printChatEntries(entries: ChatEntry[]): void {
  const sorted = [...entries].sort((a, b) => b.sortTs - a.sortTs);
  for (const e of sorted) {
    e.lines.forEach((l) => console.log(l));
    console.log();
  }
  const unreadCount = sorted.filter((e) => e.unread).length;
  console.log(
    unreadCount > 0
      ? `${c.bold}${sorted.length}${c.reset} conversations (${c.red}${c.bold}${unreadCount} unread${c.reset})`
      : `${sorted.length} conversations (0 unread)`
  );
}

export async function chatList(options: {
  pageSize?: number;
  type?: string;
  unreadOnly?: boolean;
}): Promise<void> {
  printAccountHeader();
  printChatEntries(await chatListEntries(options));
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
  amsreferences?: string[];
}

// --- Reaction helpers ---

/** Emoji map for Teams reaction keys */
const reactionEmoji: Record<string, string> = {
  like: "\u{1F44D}",
  heart: "\u{2764}\u{FE0F}",
  laugh: "\u{1F602}",
  surprised: "\u{1F62E}",
  sad: "\u{1F622}",
  angry: "\u{1F620}",
};

interface EmotionEntry {
  key: string;
  users: { mri: string; time: string; value: string }[];
}

/** Check if a message is a reaction-only message (not a regular message that happens to have reactions) */
function isReactionMessage(msg: Message): boolean {
  if (msg.messagetype === "MessageReaction") return true;
  // Some reaction notifications have messagetype "Signal/Flamingo" or contain emotion data
  if (msg.messagetype === "Signal/Flamingo") return true;
  // Check for content that is purely a reaction annotation (e.g. <e_m> tags only)
  if (msg.content && /^<e_m\b[^>]*\/>$/.test(msg.content.trim())) return true;
  // Note: messages with properties.emotions but normal messagetype (RichText/Html, Text)
  // are regular messages that have reactions ON them -- handled by emotionsSummary() instead
  return false;
}

/** Format a reaction message for display */
function formatReaction(msg: Message): string | null {
  const sender = msg.imdisplayname ?? extractUserId(msg.from) ?? "unknown";

  // Try to parse emotions from properties
  if (msg.properties?.emotions) {
    try {
      const emotions: EmotionEntry[] =
        typeof msg.properties.emotions === "string"
          ? JSON.parse(msg.properties.emotions)
          : msg.properties.emotions as EmotionEntry[];
      const reactionParts = emotions.map((e) => {
        const emoji = reactionEmoji[e.key] ?? e.key;
        return emoji;
      });
      if (reactionParts.length > 0) {
        const contentPreview = stripHtml(msg.content ?? "").slice(0, 40);
        const preview = contentPreview ? ` on: "${contentPreview}${contentPreview.length >= 40 ? "..." : ""}"` : "";
        return `${c.dim}[reaction: ${reactionParts.join(" ")}]${c.reset}${preview}`;
      }
    } catch { /* fall through */ }
  }

  // For MessageReaction type, parse content
  if (msg.messagetype === "MessageReaction" || msg.messagetype === "Signal/Flamingo") {
    // Content may contain reaction key
    const content = msg.content ?? "";
    const keyMatch = content.match(/key="([^"]+)"/);
    const emoji = keyMatch ? (reactionEmoji[keyMatch[1]] ?? keyMatch[1]) : "\u{1F44D}";
    return `${c.dim}[reaction: ${emoji}] by ${sender}${c.reset}`;
  }

  return null;
}

/** Build a summary of emotions/reactions on a message from its properties */
function emotionsSummary(msg: Message): string {
  if (!msg.properties?.emotions) return "";
  try {
    const emotions: EmotionEntry[] =
      typeof msg.properties.emotions === "string"
        ? JSON.parse(msg.properties.emotions)
        : msg.properties.emotions as EmotionEntry[];
    const parts = emotions
      .filter((e) => e.users && e.users.length > 0)
      .map((e) => {
        const emoji = reactionEmoji[e.key] ?? e.key;
        return `${emoji}${e.users.length > 1 ? e.users.length : ""}`;
      });
    if (parts.length > 0) return ` ${c.dim}${parts.join(" ")}${c.reset}`;
  } catch { /* ignore parse errors */ }
  return "";
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

  printAccountHeader();

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

    // Handle reaction messages distinctly
    if (isReactionMessage(msg)) {
      const time = formatTime(msg.originalarrivaltime ?? "");
      const reactionText = formatReaction(msg);
      if (reactionText) {
        let isNew = false;
        try { isNew = BigInt(msg.id) > myHorizon; } catch { }
        const newTag = isNew ? `${c.red}${c.bold}[NEW]${c.reset} ` : "";
        console.log(`  ${c.dim}[${time}]${c.reset} ${newTag}${reactionText}`);
      }
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

    // Show reaction summary if this message has emotions
    const reactions = emotionsSummary(msg);

    console.log(`${indent}${msgIdTag}${c.dim}[${time}]${c.reset} ${newTag}${replyPrefix}${c.green}${c.bold}${sender}${c.reset}: ${content}${reactions}${threadTag}`);
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

  printAccountHeader();

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

    // Handle reaction messages distinctly in threads too
    if (isReactionMessage(msg)) {
      const time = formatTime(msg.originalarrivaltime ?? "");
      const reactionText = formatReaction(msg);
      if (reactionText) {
        console.log(`  ${c.dim}[${time}]${c.reset} ${reactionText}`);
      }
      continue;
    }

    const sender = msg.imdisplayname ?? extractUserId(msg.from) ?? "system";
    const time = formatTime(msg.originalarrivaltime ?? "");
    const content = stripHtml(msg.content ?? "");
    const isRoot = msg.id === rootMessageId;
    const indent = isRoot ? "" : "  ";

    // Show reaction summary if this message has emotions
    const reactions = emotionsSummary(msg);

    console.log(`${indent}${c.dim}[${time}]${c.reset} ${c.green}${c.bold}${sender}${c.reset}: ${content}${reactions}`);
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
