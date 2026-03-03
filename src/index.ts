import { program } from "commander";
import { tokenStatus, refresh, isTokenValid, tryRefresh, deviceCodeLogin } from "./auth.js";
import { chatList, chatRead, chatSend, chatMarkRead, chatThread } from "./api.js";
import { mailList, mailRead, mailSearch, mailDraft, mailSend, mailCompose, mailReply, mailOpen, mailAttachments, calendarList, calendarRead, calendarToday, calendarSchedule, calendarFindSlot } from "./outlook-api.js";
import { requireTouchId } from "./touchid.js";
import { resolveId } from "./id-map.js";

/** Ensure we have a valid token, auto-refreshing or prompting device login */
async function ensureToken(): Promise<void> {
  if (isTokenValid()) return;

  console.error("Token expired. Trying refresh token...");
  if (await tryRefresh(true)) return;

  console.error("Refresh failed. Run: ms-cli auth login");
  process.exit(1);
}

program.name("ms-cli").description("Teams Internal API CLI").version("0.1.0");

// --- auth ---
const auth = program.command("auth").description("Token management");

auth
  .command("login")
  .description("Login via device code flow")
  .action(async () => {
    await deviceCodeLogin();
  });

auth
  .command("status")
  .description("Show token expiration info")
  .action(() => {
    tokenStatus();
  });

auth
  .command("refresh")
  .description("Refresh skypetoken using saved refresh token")
  .action(async () => {
    await refresh();
  });

// --- chat ---
const chat = program.command("chat").description("Chat operations");

chat
  .command("list")
  .description("List conversations")
  .option("-n, --page-size <n>", "Number of conversations", "20")
  .option("-t, --type <type>", "Filter: chat, channel, meeting")
  .option("-u, --unread", "Show only unread conversations")
  .action(async (opts) => {
    await ensureToken();
    await chatList({ pageSize: parseInt(opts.pageSize), type: opts.type, unreadOnly: opts.unread });
  });

chat
  .command("read <id>")
  .description("Read messages from a conversation")
  .option("-n, --limit <n>", "Number of messages", "20")
  .option("--json", "Output raw JSON")
  .action(async (id: string, opts) => {
    await ensureToken();
    await chatRead(id, { limit: parseInt(opts.limit), json: opts.json });
  });

chat
  .command("send <id> <message>")
  .description("Send a message to a conversation (requires Touch ID)")
  .action(async (id: string, message: string) => {
    requireTouchId("ms-cli: Teams メッセージ送信");
    await ensureToken();
    await chatSend(id, message);
  });

chat
  .command("thread <conversationId> <messageId>")
  .description("Read a thread (reply chain) in a channel")
  .option("-n, --limit <n>", "Max messages to scan", "200")
  .option("--json", "Output raw JSON")
  .action(async (conversationId: string, messageId: string, opts) => {
    await ensureToken();
    await chatThread(conversationId, resolveId(messageId), { limit: parseInt(opts.limit), json: opts.json });
  });

chat
  .command("mark-read <id>")
  .description("Mark a conversation as read")
  .action(async (id: string) => {
    await ensureToken();
    await chatMarkRead(id);
  });

// --- mail ---
const mail = program.command("mail").description("Outlook mail operations");

mail
  .command("list")
  .description("List inbox messages")
  .option("-n, --page-size <n>", "Number of messages", "15")
  .option("-u, --unread", "Show only unread messages")
  .option("-f, --folder <folder>", "Mail folder (default: inbox)")
  .action(async (opts) => {
    await mailList({
      pageSize: parseInt(opts.pageSize),
      unreadOnly: opts.unread,
      folder: opts.folder,
    });
  });

mail
  .command("read <id>")
  .description("Read a specific email message")
  .option("--json", "Output raw JSON")
  .action(async (id: string, opts) => {
    await mailRead(resolveId(id), { json: opts.json });
  });

mail
  .command("search <query>")
  .description("Search emails")
  .option("-n, --page-size <n>", "Number of results", "10")
  .action(async (query: string, opts) => {
    await mailSearch(query, { pageSize: parseInt(opts.pageSize) });
  });

mail
  .command("draft")
  .description("Create a draft email")
  .requiredOption("--to <addrs...>", "Recipients (comma-separated or multiple)")
  .requiredOption("-s, --subject <subject>", "Subject")
  .requiredOption("-b, --body <body>", "Body text")
  .option("--cc <addrs...>", "CC recipients")
  .option("--html", "Body is HTML")
  .option("--importance <level>", "Normal, Low, or High")
  .action(async (opts) => {
    await mailDraft({
      to: opts.to,
      subject: opts.subject,
      body: opts.body,
      cc: opts.cc,
      html: opts.html,
      importance: opts.importance,
    });
  });

mail
  .command("send <id>")
  .description("Send a draft email by message ID (requires Touch ID)")
  .action(async (id: string) => {
    requireTouchId("ms-cli: メール送信");
    await mailSend(resolveId(id));
  });

mail
  .command("attachments <id>")
  .description("List or download attachments from an email")
  .option("-l, --list", "List attachments without downloading")
  .option("-o, --out-dir <dir>", "Output directory (default: current dir)")
  .action(async (id: string, opts) => {
    await mailAttachments(resolveId(id), { list: opts.list, outDir: opts.outDir });
  });

mail
  .command("open <id>")
  .description("Open a message in Outlook Web (browser)")
  .action(async (id: string) => {
    await mailOpen(resolveId(id));
  });

mail
  .command("reply <id>")
  .description("Create a reply draft (reply-all by default)")
  .requiredOption("-b, --body <body>", "Reply body text")
  .option("--no-all", "Reply to sender only (default: reply all)")
  .action(async (id: string, opts) => {
    await mailReply(resolveId(id), { body: opts.body, all: opts.all });
  });

mail
  .command("compose")
  .description("Compose and send an email immediately (requires Touch ID)")
  .requiredOption("--to <addrs...>", "Recipients")
  .requiredOption("-s, --subject <subject>", "Subject")
  .requiredOption("-b, --body <body>", "Body text")
  .option("--cc <addrs...>", "CC recipients")
  .option("--html", "Body is HTML")
  .option("--importance <level>", "Normal, Low, or High")
  .action(async (opts) => {
    requireTouchId("ms-cli: メール作成・送信");
    await mailCompose({
      to: opts.to,
      subject: opts.subject,
      body: opts.body,
      cc: opts.cc,
      html: opts.html,
      importance: opts.importance,
    });
  });

// --- calendar ---
const cal = program.command("cal").description("Calendar operations");

cal
  .command("list")
  .description("List upcoming events")
  .option("-d, --days <n>", "Number of days ahead", "7")
  .option("-n, --page-size <n>", "Max events", "30")
  .action(async (opts) => {
    await calendarList({ days: parseInt(opts.days), pageSize: parseInt(opts.pageSize) });
  });

cal
  .command("today")
  .description("Show today's schedule")
  .action(async () => {
    await calendarToday();
  });

cal
  .command("read <id>")
  .description("Show event details")
  .option("--json", "Output raw JSON")
  .action(async (id: string, opts) => {
    await calendarRead(id, { json: opts.json });
  });

cal
  .command("schedule <emails...>")
  .description("Show schedule for users on a specific date")
  .option("-d, --date <date>", "Date (YYYY-MM-DD)", new Date().toLocaleDateString("sv-SE", { timeZone: "Asia/Tokyo" }))
  .option("--start-hour <n>", "Start hour", "9")
  .option("--end-hour <n>", "End hour", "17")
  .action(async (emails: string[], opts) => {
    await calendarSchedule({
      emails,
      date: opts.date,
      startHour: parseInt(opts.startHour),
      endHour: parseInt(opts.endHour),
    });
  });

cal
  .command("find-slot <emails...>")
  .description("Find common free slots across users")
  .option("--duration <min>", "Meeting duration in minutes", "60")
  .option("--days <n>", "Number of weekdays to search", "5")
  .option("--start-hour <n>", "Start hour", "9")
  .option("--end-hour <n>", "End hour", "17")
  .option("--start-date <date>", "Start searching from (YYYY-MM-DD)")
  .action(async (emails: string[], opts) => {
    await calendarFindSlot({
      emails,
      duration: parseInt(opts.duration),
      days: parseInt(opts.days),
      startHour: parseInt(opts.startHour),
      endHour: parseInt(opts.endHour),
      startDate: opts.startDate,
    });
  });

program.parse();
