import { program } from "commander";
import { tokenStatus, refresh, isTokenValid, tryRefresh, deviceCodeLogin, authList, authUse, authRemove, syncTenants } from "./auth.js";
import { listAccounts } from "./config.js";
import { chatList, chatRead, chatSend, chatMarkRead, chatThread, chatListEntries, printChatEntries, type ChatEntry } from "./api.js";
import { mailList, mailRead, mailSearch, mailDraft, mailSend, mailCompose, mailReply, mailOpen, mailAttachments, calendarList, calendarRead, calendarToday, calendarSchedule, calendarFindSlot } from "./outlook-api.js";
import { spSites, spDrives, spFiles, spDownload, spSearch, spRecent, spOpen, spUpload, spConvert, spDelete, spUploadConvert } from "./sharepoint-api.js";
import { formsList, formsRead, formsResponses, formsGroup } from "./forms-api.js";
import { requireTouchId } from "./touchid.js";
import { resolveId } from "./id-map.js";

// Surface expected errors (e.g. unknown account) as clean messages, not stack traces.
process.on("uncaughtException", (e) => {
  console.error((e as Error).message ?? e);
  process.exit(1);
});
process.on("unhandledRejection", (e) => {
  console.error((e as Error)?.message ?? e);
  process.exit(1);
});

/** Temporarily pin the active account to `key` while running fn. */
async function withAccount<T>(key: string, fn: () => Promise<T>): Promise<T> {
  const prev = process.env.MS_CLI_ACCOUNT;
  process.env.MS_CLI_ACCOUNT = key;
  try {
    return await fn();
  } finally {
    if (prev === undefined) delete process.env.MS_CLI_ACCOUNT;
    else process.env.MS_CLI_ACCOUNT = prev;
  }
}

/** Short tenant label for inline tagging, e.g. "丸紅" / "Digital Experts". */
function shortLabel(cfg: { tenantName?: string; tenantId?: string }): string {
  const n = (cfg.tenantName ?? "").split(/株式会社|（|\(/)[0].trim();
  return n || cfg.tenantId?.slice(0, 8) || "?";
}

/** True for errors that just mean "this account has no mailbox here" (guests). */
function isNoMailboxError(e: unknown): boolean {
  const m = (e as Error)?.message ?? "";
  return m.includes("MailboxNotEnabledForRESTAPI") || m.includes("error 404");
}

/**
 * Run a read-many operation across accounts (or just `accountRef` if given),
 * so the user never has to switch tenants. Each call labels its own output.
 * `membersOnly` skips guest tenants (used for mail/calendar, which need a mailbox).
 */
async function fanOut(
  accountRef: string | undefined,
  fn: () => Promise<void>,
  opts: { membersOnly?: boolean } = {}
): Promise<void> {
  let accounts = accountRef
    ? listAccounts().filter((a) => a.key === accountRef)
    : listAccounts();
  if (!accountRef && opts.membersOnly) {
    accounts = accounts.filter((a) => a.config.userType !== "guest");
  }
  const keys = accounts.length ? accounts.map((a) => a.key) : accountRef ? [accountRef] : [];
  if (keys.length <= 1) {
    await (keys[0] ? withAccount(keys[0], fn) : fn());
    return;
  }
  for (let i = 0; i < keys.length; i++) {
    if (i > 0) console.log();
    try {
      await withAccount(keys[i], fn);
    } catch (e) {
      if (isNoMailboxError(e)) continue; // guest tenant w/o mailbox — silently skip
      console.error(`[${keys[i]}] ${(e as Error).message}`);
    }
  }
}

/** Merge chat conversations from all tenants into one newest-first list. */
async function mergedChatList(
  accountRef: string | undefined,
  options: { pageSize?: number; type?: string; unreadOnly?: boolean }
): Promise<void> {
  if (accountRef) {
    await withAccount(accountRef, async () => {
      await ensureToken();
      await chatList(options);
    });
    return;
  }
  const accounts = listAccounts();
  if (accounts.length <= 1) {
    await ensureToken();
    await chatList(options);
    return;
  }
  const all: ChatEntry[] = [];
  for (const a of accounts) {
    try {
      await withAccount(a.key, async () => {
        await ensureToken();
        all.push(...(await chatListEntries({ ...options, tag: shortLabel(a.config) })));
      });
    } catch (e) {
      console.error(`[${shortLabel(a.config)}] ${(e as Error).message}`);
    }
  }
  all.sort((x, y) => y.sortTs - x.sortTs);
  printChatEntries(options.pageSize ? all.slice(0, options.pageSize) : all);
}

/**
 * Run a single-resource lookup against each account until one succeeds, so an
 * ID can be read without knowing which tenant it belongs to.
 */
async function firstAccount(accountRef: string | undefined, fn: () => Promise<void>): Promise<void> {
  const keys = accountRef ? [accountRef] : listAccounts().map((a) => a.key);
  if (keys.length <= 1) {
    await (keys[0] ? withAccount(keys[0], fn) : fn());
    return;
  }
  let lastErr: unknown;
  for (const key of keys) {
    try {
      await withAccount(key, fn);
      return;
    } catch (e) {
      lastErr = e;
    }
  }
  console.error(lastErr ? (lastErr as Error).message : "Not found in any tenant.");
  process.exit(1);
}

/** Ensure we have a valid token, auto-refreshing or prompting device login */
async function ensureToken(): Promise<void> {
  if (isTokenValid()) return;

  console.error("Token expired. Trying refresh token...");
  if (await tryRefresh(true)) return;

  console.error("Refresh failed. Run: ms-cli auth login");
  process.exit(1);
}

program.name("ms-cli").description("Teams Internal API CLI").version("0.3.0");

// --- auth ---
const auth = program.command("auth").description("Token management");

auth
  .command("login")
  .description("Login via device code flow (adds/switches account)")
  .option("--name <name>", "Account label (default: your email/UPN)")
  .option("--tenant <id>", "Tenant ID or domain to log into (default: common)")
  .action(async (opts) => {
    await deviceCodeLogin({ name: opts.name, tenant: opts.tenant });
  });

auth
  .command("status")
  .description("Show token expiration info for the active account")
  .action(() => {
    tokenStatus();
  });

auth
  .command("list")
  .description("List stored accounts (* = active)")
  .action(() => {
    authList();
  });

auth
  .command("sync")
  .description("Discover all tenants for the current login and register them")
  .action(async () => {
    await syncTenants();
  });

auth
  .command("use <key>")
  .description("Switch the active account")
  .action((key: string) => {
    authUse(key);
  });

auth
  .command("remove <key>")
  .description("Remove a stored account")
  .action((key: string) => {
    authRemove(key);
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
  .option("-a, --account <ref>", "Limit to one account (default: all tenants)")
  .action(async (opts) => {
    await mergedChatList(opts.account, {
      pageSize: parseInt(opts.pageSize),
      type: opts.type,
      unreadOnly: opts.unread,
    });
  });

chat
  .command("read <id>")
  .description("Read messages from a conversation")
  .option("-n, --limit <n>", "Number of messages", "20")
  .option("--json", "Output raw JSON")
  .option("-a, --account <ref>", "Limit to one account (default: search all tenants)")
  .action(async (id: string, opts) => {
    await firstAccount(opts.account, async () => {
      await ensureToken();
      await chatRead(id, { limit: parseInt(opts.limit), json: opts.json });
    });
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
  .option("-a, --account <ref>", "Limit to one account (default: search all tenants)")
  .action(async (conversationId: string, messageId: string, opts) => {
    await firstAccount(opts.account, async () => {
      await ensureToken();
      await chatThread(conversationId, resolveId(messageId), { limit: parseInt(opts.limit), json: opts.json });
    });
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
  .option("-a, --account <ref>", "Limit to one account (default: all tenants)")
  .action(async (opts) => {
    await fanOut(opts.account, async () => {
      await mailList({
        pageSize: parseInt(opts.pageSize),
        unreadOnly: opts.unread,
        folder: opts.folder,
      });
    }, { membersOnly: true });
  });

mail
  .command("read <id>")
  .description("Read a specific email message")
  .option("--json", "Output raw JSON")
  .option("-a, --account <ref>", "Limit to one account (default: search all tenants)")
  .action(async (id: string, opts) => {
    await firstAccount(opts.account, async () => {
      await mailRead(resolveId(id), { json: opts.json });
    });
  });

mail
  .command("search <query>")
  .description("Search emails")
  .option("-n, --page-size <n>", "Number of results", "10")
  .option("-a, --account <ref>", "Limit to one account (default: all tenants)")
  .action(async (query: string, opts) => {
    await fanOut(opts.account, async () => {
      await mailSearch(query, { pageSize: parseInt(opts.pageSize) });
    }, { membersOnly: true });
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
  .option("-u, --user <email>", "View another user's calendar")
  .option("-a, --account <ref>", "Limit to one account (default: all tenants)")
  .action(async (opts) => {
    await fanOut(opts.account, async () => {
      await calendarList({ days: parseInt(opts.days), pageSize: parseInt(opts.pageSize), user: opts.user });
    }, { membersOnly: true });
  });

cal
  .command("today")
  .description("Show today's schedule")
  .option("-u, --user <email>", "View another user's calendar")
  .option("-a, --account <ref>", "Limit to one account (default: all tenants)")
  .action(async (opts) => {
    await fanOut(opts.account, async () => {
      await calendarToday({ user: opts.user });
    }, { membersOnly: true });
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

// --- forms ---
const forms = program.command("forms").description("Microsoft Forms (unofficial API)");

forms
  .command("list")
  .description("List your forms")
  .option("-n, --page-size <n>", "Number of forms", "20")
  .option("--json", "Output raw JSON")
  .action(async (opts) => {
    await ensureToken();
    await formsList({ pageSize: parseInt(opts.pageSize), json: opts.json });
  });

forms
  .command("read <formIdOrUrl>")
  .description("Show form details and questions (accepts form ID or URL)")
  .option("--owner <ownerId>", "Owner user ID (auto-resolved if omitted)")
  .option("--json", "Output raw JSON")
  .action(async (formIdOrUrl: string, opts) => {
    await ensureToken();
    await formsRead(formIdOrUrl, { owner: opts.owner, json: opts.json });
  });

forms
  .command("responses <formIdOrUrl>")
  .description("List form responses (accepts form ID or URL)")
  .option("--owner <ownerId>", "Owner user ID (auto-resolved if omitted)")
  .option("-n, --page-size <n>", "Number of responses", "20")
  .option("--json", "Output raw JSON")
  .action(async (formIdOrUrl: string, opts) => {
    await ensureToken();
    await formsResponses(formIdOrUrl, { owner: opts.owner, pageSize: parseInt(opts.pageSize), json: opts.json });
  });

forms
  .command("group <groupId>")
  .description("List forms in a Microsoft 365 group")
  .option("-n, --page-size <n>", "Number of forms", "20")
  .option("--json", "Output raw JSON")
  .action(async (groupId: string, opts) => {
    await ensureToken();
    await formsGroup(groupId, { pageSize: parseInt(opts.pageSize), json: opts.json });
  });

// --- sharepoint ---
const sp = program.command("sp").description("SharePoint operations");

sp
  .command("sites")
  .description("Search / list SharePoint sites")
  .option("-q, --query <query>", "Search query (default: all)")
  .option("-n, --page-size <n>", "Number of results", "20")
  .action(async (opts) => {
    await ensureToken();
    await spSites({ query: opts.query, pageSize: parseInt(opts.pageSize) });
  });

sp
  .command("drives <siteId>")
  .description("List document libraries (drives) in a site")
  .action(async (siteId: string) => {
    await ensureToken();
    await spDrives(siteId);
  });

sp
  .command("files <driveId>")
  .description("List files in a drive")
  .option("-p, --path <path>", "Folder path within drive")
  .option("-n, --page-size <n>", "Number of items", "30")
  .option("--json", "Output raw JSON")
  .action(async (driveId: string, opts) => {
    await ensureToken();
    await spFiles(driveId, { path: opts.path, pageSize: parseInt(opts.pageSize), json: opts.json });
  });

sp
  .command("download <driveId> <itemId>")
  .description("Download a file")
  .option("-o, --out-dir <dir>", "Output directory (default: current dir)")
  .action(async (driveId: string, itemId: string, opts) => {
    await ensureToken();
    await spDownload(driveId, itemId, { outDir: opts.outDir });
  });

sp
  .command("search <query>")
  .description("Search files across SharePoint")
  .option("-n, --page-size <n>", "Number of results", "15")
  .action(async (query: string, opts) => {
    await ensureToken();
    await spSearch(query, { pageSize: parseInt(opts.pageSize) });
  });

sp
  .command("recent")
  .description("Show recently accessed files")
  .option("-n, --page-size <n>", "Number of files", "20")
  .action(async (opts) => {
    await ensureToken();
    await spRecent({ pageSize: parseInt(opts.pageSize) });
  });

sp
  .command("open <driveId> <itemId>")
  .description("Open a file in the browser")
  .action(async (driveId: string, itemId: string) => {
    await ensureToken();
    await spOpen(driveId, itemId);
  });

sp
  .command("upload <driveId> <localPath>")
  .description("Upload a local file to a drive")
  .option("-p, --remote-path <path>", "Remote filename/path (default: same as local)")
  .action(async (driveId: string, localPath: string, opts) => {
    await ensureToken();
    await spUpload(driveId, localPath, { remotePath: opts.remotePath });
  });

sp
  .command("convert <driveId> <itemId>")
  .description("Download a file converted to another format (e.g. pptx→pdf)")
  .option("-f, --format <fmt>", "Target format: pdf, html, jpg, png, glb", "pdf")
  .option("-o, --out-dir <dir>", "Output directory (default: current dir)")
  .action(async (driveId: string, itemId: string, opts) => {
    await ensureToken();
    await spConvert(driveId, itemId, { format: opts.format, outDir: opts.outDir });
  });

sp
  .command("delete <driveId> <itemId>")
  .description("Delete a file from a drive")
  .action(async (driveId: string, itemId: string) => {
    await ensureToken();
    await spDelete(driveId, itemId);
  });

sp
  .command("to-pdf <driveId> <localPath>")
  .description("Upload a file, convert to PDF, download, and delete remote (one-shot)")
  .option("-f, --format <fmt>", "Target format", "pdf")
  .option("-o, --out-dir <dir>", "Output directory (default: current dir)")
  .action(async (driveId: string, localPath: string, opts) => {
    await ensureToken();
    await spUploadConvert(driveId, localPath, { format: opts.format, outDir: opts.outDir });
  });

// --- update ---
program
  .command("update")
  .description("Self-update to the latest release")
  .option("--path <path>", "Path to the binary to update (default: auto-detect)")
  .action(async (opts) => {
    const { realpathSync, writeFileSync, renameSync, chmodSync } = await import("fs");
    const arch = process.arch === "arm64" ? "arm64" : "x64";
    const asset = `ms-cli-darwin-${arch}`;
    const url = `https://github.com/Mojashi/ms-cli/releases/latest/download/${asset}`;

    // Detect the binary path: explicit flag > argv[1] for bun-compiled > which ms-cli
    let self: string;
    if (opts.path) {
      self = realpathSync(opts.path);
    } else {
      // For bun-compiled binaries, execPath IS the binary
      // For tsx/node, we need to find the actual installed binary
      const { execSync } = await import("child_process");
      try {
        self = realpathSync(execSync("which ms-cli", { encoding: "utf-8" }).trim());
      } catch {
        console.error("Cannot detect ms-cli binary path. Use --path to specify.");
        process.exit(1);
      }
    }

    console.log(`Downloading ${asset}...`);
    const res = await fetch(url, { redirect: "follow" });
    if (!res.ok) {
      console.error(`Download failed: ${res.status}`);
      process.exit(1);
    }

    const tmpPath = `${self}.tmp`;
    writeFileSync(tmpPath, Buffer.from(await res.arrayBuffer()));
    chmodSync(tmpPath, 0o755);
    renameSync(tmpPath, self);
    console.log(`Updated: ${self}`);
  });

program.parse();
