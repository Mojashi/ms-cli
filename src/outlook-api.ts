import { execSync } from "child_process";
import { writeFileSync, mkdirSync } from "fs";
import { join } from "path";
import { loadConfig, saveConfig } from "./config.js";
import { shortId, registerIds } from "./id-map.js";
import { ensureGraphToken } from "./sharepoint-api.js";
import { printAccountHeader } from "./api.js";

const TEAMS_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"; // Microsoft Teams native client
const OUTLOOK_BASE = "https://outlook.office.com/api/v2.0/me";

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
};

async function ensureOutlookToken(): Promise<string> {
  const config = loadConfig();

  // Check if we have a valid outlook token
  if (config.outlookToken) {
    try {
      const payload = JSON.parse(
        Buffer.from(config.outlookToken.split(".")[1] + "==", "base64url").toString()
      );
      if (payload.exp && payload.exp > Date.now() / 1000 + 60) {
        return config.outlookToken;
      }
    } catch {}
  }

  // Need to refresh
  if (!config.refreshToken) {
    console.error("No refresh token. Run: ms-cli auth login");
    process.exit(1);
  }

  console.error("Refreshing Outlook token...");
  const tenantId = config.tenantId ?? "common";
  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Origin: "https://teams.microsoft.com",
      },
      body: new URLSearchParams({
        client_id: config.clientId ?? TEAMS_CLIENT_ID,
        grant_type: "refresh_token",
        refresh_token: config.refreshToken,
        scope: "https://outlook.office.com/.default openid profile offline_access",
      }),
    }
  );

  if (!res.ok) {
    const text = await res.text();
    console.error(`Token refresh failed (${res.status}): ${text.slice(0, 200)}`);
    process.exit(1);
  }

  const data = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
  };

  config.outlookToken = data.access_token;
  if (data.refresh_token) {
    config.refreshToken = data.refresh_token;
    config.refreshTokenIssuedAt = Math.floor(Date.now() / 1000);
  }
  saveConfig(config);
  console.error("Outlook token refreshed.");
  return data.access_token;
}

async function outlookGet(path: string): Promise<unknown> {
  const token = await ensureOutlookToken();
  const url = path.startsWith("http") ? path : `${OUTLOOK_BASE}${path}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Outlook API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

async function outlookPost(path: string, body?: unknown): Promise<unknown> {
  const token = await ensureOutlookToken();
  const url = path.startsWith("http") ? path : `${OUTLOOK_BASE}${path}`;
  const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
  const init: RequestInit = { method: "POST", headers };
  if (body) {
    headers["Content-Type"] = "application/json";
    init.body = JSON.stringify(body);
  }
  const res = await fetch(url, init);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Outlook API error ${res.status}: ${text.slice(0, 300)}`);
  }
  if (res.status === 202 || res.headers.get("content-length") === "0") return {};
  return res.json();
}

// --- Graph API helpers ---

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

async function graphGet(path: string): Promise<unknown> {
  const token = await ensureGraphToken();
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

async function graphPost(path: string, body?: unknown): Promise<unknown> {
  const token = await ensureGraphToken();
  const url = path.startsWith("http") ? path : `${GRAPH_BASE}${path}`;
  const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
  const init: RequestInit = { method: "POST", headers };
  if (body) {
    headers["Content-Type"] = "application/json";
    init.body = JSON.stringify(body);
  }
  const res = await fetch(url, init);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API error ${res.status}: ${text.slice(0, 300)}`);
  }
  if (res.status === 202 || res.headers.get("content-length") === "0") return {};
  return res.json();
}

// --- Mail List ---

interface ODataResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

interface MailMessage {
  Id: string;
  Subject: string;
  From: { EmailAddress: { Name: string; Address: string } };
  ReceivedDateTime: string;
  IsRead: boolean;
  BodyPreview?: string;
  HasAttachments?: boolean;
  Importance?: string;
  ToRecipients?: Array<{ EmailAddress: { Name: string; Address: string } }>;
  Body?: { ContentType: string; Content: string };
}

export async function mailList(options: {
  pageSize?: number;
  unreadOnly?: boolean;
  folder?: string;
}): Promise<void> {
  const pageSize = options.pageSize ?? 15;
  const folder = options.folder ?? "inbox";
  const filter = options.unreadOnly ? "&$filter=IsRead eq false" : "";

  const data = (await outlookGet(
    `/mailfolders/${folder}/messages?$top=${pageSize}&$select=Subject,From,ReceivedDateTime,IsRead,BodyPreview,HasAttachments,Importance&$orderby=ReceivedDateTime desc${filter}`
  )) as ODataResponse<MailMessage>;

  printAccountHeader();
  registerIds(data.value.map((m) => m.Id));

  for (const mail of data.value) {
    const sid = shortId(mail.Id);
    const from = mail.From?.EmailAddress?.Name ?? mail.From?.EmailAddress?.Address ?? "unknown";
    const time = formatTime(mail.ReceivedDateTime);
    const subject = mail.Subject ?? "(no subject)";
    const preview = (mail.BodyPreview ?? "").replace(/\r?\n/g, " ").slice(0, 70);
    const unread = !mail.IsRead;
    const attach = mail.HasAttachments ? " [+]" : "";
    const imp = mail.Importance === "High" ? ` ${c.red}[!]${c.reset}` : "";

    const marker = unread
      ? `${c.bgRed}${c.white}${c.bold} NEW ${c.reset} `
      : "      ";
    const subjectStyle = unread ? `${c.bold}${c.white}` : `${c.dim}`;

    console.log(`${marker}${c.yellow}${sid}${c.reset} ${subjectStyle}${subject}${c.reset}${attach}${imp}`);
    console.log(`        ${c.green}${from}${c.reset} ${c.dim}(${time})${c.reset}`);
    if (preview) {
      console.log(`        ${c.dim}${preview}${c.reset}`);
    }
    console.log();
  }

  const unreadCount = data.value.filter((m) => !m.IsRead).length;
  console.log(
    `${c.bold}${data.value.length}${c.reset} messages` +
      (unreadCount > 0 ? ` (${c.red}${c.bold}${unreadCount} unread${c.reset})` : "")
  );
}

// --- Mail Read ---

export async function mailRead(
  messageId: string,
  options: { json?: boolean }
): Promise<void> {
  const data = (await outlookGet(
    `/messages/${messageId}?$select=Subject,From,ToRecipients,ReceivedDateTime,Body,IsRead,HasAttachments,Importance`
  )) as MailMessage;

  if (options.json) {
    console.log(JSON.stringify(data, null, 2));
    return;
  }

  const from = data.From?.EmailAddress
    ? `${data.From.EmailAddress.Name} <${data.From.EmailAddress.Address}>`
    : "unknown";
  const to = (data.ToRecipients ?? [])
    .map((r) => r.EmailAddress?.Name ?? r.EmailAddress?.Address)
    .join(", ");

  console.log(`${c.bold}Subject:${c.reset} ${data.Subject}`);
  console.log(`${c.bold}From:${c.reset}    ${c.green}${from}${c.reset}`);
  console.log(`${c.bold}To:${c.reset}      ${to}`);
  console.log(`${c.bold}Date:${c.reset}    ${formatTime(data.ReceivedDateTime)}`);
  if (data.HasAttachments) {
    const attachments = await listAttachments(messageId);
    const names = attachments.map((a) => a.Name).join(", ");
    console.log(`${c.bold}Attach:${c.reset}  ${c.yellow}${names}${c.reset}`);
  }
  console.log(`${c.dim}${"─".repeat(60)}${c.reset}`);

  if (data.Body) {
    const content =
      data.Body.ContentType === "HTML"
        ? stripHtml(data.Body.Content)
        : data.Body.Content;
    console.log(content);
  }
}

// --- Mail Attachments ---

interface Attachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  ContentBytes?: string;
  IsInline?: boolean;
}

async function listAttachments(messageId: string): Promise<Attachment[]> {
  const data = (await outlookGet(
    `/messages/${messageId}/attachments?$select=Id,Name,ContentType,Size,IsInline`
  )) as ODataResponse<Attachment>;
  return data.value;
}

export async function mailAttachments(
  messageId: string,
  options: { outDir?: string; list?: boolean }
): Promise<void> {
  const attachments = await listAttachments(messageId);

  if (attachments.length === 0) {
    console.log("No attachments.");
    return;
  }

  if (options.list) {
    console.log(`${c.bold}${c.cyan}── Attachments ──${c.reset}\n`);
    for (const a of attachments) {
      const sizeStr = a.Size > 1024 * 1024
        ? `${(a.Size / 1024 / 1024).toFixed(1)} MB`
        : `${Math.ceil(a.Size / 1024)} KB`;
      const inline = a.IsInline ? ` ${c.dim}(inline)${c.reset}` : "";
      console.log(`  ${c.yellow}${a.Name}${c.reset}  ${c.dim}${sizeStr}  ${a.ContentType}${c.reset}${inline}`);
    }
    console.log(`\n${c.bold}${attachments.length}${c.reset} attachment(s)`);
    return;
  }

  // Download attachments
  const outDir = options.outDir ?? ".";
  mkdirSync(outDir, { recursive: true });

  for (const a of attachments) {
    const full = (await outlookGet(
      `/messages/${messageId}/attachments/${a.Id}`
    )) as Attachment;

    if (!full.ContentBytes) {
      console.log(`${c.yellow}Skipped:${c.reset} ${a.Name} (no content, may be a reference attachment)`);
      continue;
    }

    const buf = Buffer.from(full.ContentBytes, "base64");
    const outPath = join(outDir, a.Name);
    writeFileSync(outPath, buf);

    const sizeStr = buf.length > 1024 * 1024
      ? `${(buf.length / 1024 / 1024).toFixed(1)} MB`
      : `${Math.ceil(buf.length / 1024)} KB`;
    console.log(`${c.green}Saved:${c.reset} ${outPath} ${c.dim}(${sizeStr})${c.reset}`);
  }
}

// --- Mail Search ---

export async function mailSearch(
  query: string,
  options: { pageSize?: number }
): Promise<void> {
  const pageSize = options.pageSize ?? 10;

  const data = (await outlookGet(
    `/messages?$search="${encodeURIComponent(query)}"&$top=${pageSize}&$select=Subject,From,ReceivedDateTime,IsRead,BodyPreview`
  )) as ODataResponse<MailMessage>;

  printAccountHeader();
  if (data.value.length === 0) {
    console.log("No results found.");
    return;
  }

  registerIds(data.value.map((m) => m.Id));

  for (const mail of data.value) {
    const sid = shortId(mail.Id);
    const from = mail.From?.EmailAddress?.Name ?? mail.From?.EmailAddress?.Address ?? "unknown";
    const time = formatTime(mail.ReceivedDateTime);
    const subject = mail.Subject ?? "(no subject)";
    const unread = !mail.IsRead;
    const preview = (mail.BodyPreview ?? "").replace(/\r?\n/g, " ").slice(0, 70);
    const marker = unread ? `${c.red}*${c.reset}` : " ";

    console.log(`${marker} ${c.yellow}${sid}${c.reset} ${c.bold}${subject}${c.reset}`);
    console.log(`  ${c.green}${from}${c.reset} ${c.dim}(${time})${c.reset}`);
    console.log(`  ${c.dim}${preview}${c.reset}`);
    console.log();
  }

  console.log(`${data.value.length} results`);
}

// --- Mail Draft ---

export async function mailDraft(options: {
  subject: string;
  body: string;
  to: string[];
  cc?: string[];
  importance?: string;
  html?: boolean;
}): Promise<void> {
  const toRecipients = options.to.map(parseRecipient);
  const ccRecipients = (options.cc ?? []).map(parseRecipient);

  const body = prepareBody(options.body, options.html);
  const payload: Record<string, unknown> = {
    Subject: options.subject,
    Body: body,
    ToRecipients: toRecipients,
  };
  if (ccRecipients.length > 0) payload.CcRecipients = ccRecipients;
  if (options.importance) payload.Importance = options.importance;

  const data = (await outlookPost("/messages", payload)) as MailMessage;

  console.log(`${c.green}Draft created.${c.reset}`);
  console.log(`${c.bold}Subject:${c.reset} ${data.Subject}`);
  console.log(`${c.bold}To:${c.reset}      ${options.to.join(", ")}`);
  if (options.cc?.length) console.log(`${c.bold}Cc:${c.reset}      ${options.cc.join(", ")}`);
  const sid = shortId(data.Id);
  registerIds([data.Id]);
  console.log(`${c.bold}id:${c.reset}      ${c.yellow}${sid}${c.reset}`);
  console.log(`\nPreview: ${c.cyan}ms-cli mail open ${sid}${c.reset}`);
  console.log(`Send:    ${c.cyan}ms-cli mail send ${sid}${c.reset}`);
}

// --- Mail Send ---

export async function mailSend(messageId: string): Promise<void> {
  await outlookPost(`/messages/${messageId}/send`);
  console.log(`${c.green}Message sent.${c.reset}`);
}

// --- Mail Draft+Send (compose and send immediately) ---

export async function mailCompose(options: {
  subject: string;
  body: string;
  to: string[];
  cc?: string[];
  importance?: string;
  html?: boolean;
}): Promise<void> {
  const toRecipients = options.to.map(parseRecipient);
  const ccRecipients = (options.cc ?? []).map(parseRecipient);

  const body = prepareBody(options.body, options.html);
  const payload: Record<string, unknown> = {
    Subject: options.subject,
    Body: body,
    ToRecipients: toRecipients,
  };
  if (ccRecipients.length > 0) payload.CcRecipients = ccRecipients;
  if (options.importance) payload.Importance = options.importance;

  await outlookPost("/sendmail", { Message: payload });
  console.log(`${c.green}Message sent.${c.reset}`);
  console.log(`${c.bold}Subject:${c.reset} ${options.subject}`);
  console.log(`${c.bold}To:${c.reset}      ${options.to.join(", ")}`);
}

// --- Mail Reply Draft ---

async function outlookPatch(path: string, body: unknown): Promise<unknown> {
  const token = await ensureOutlookToken();
  const url = path.startsWith("http") ? path : `${OUTLOOK_BASE}${path}`;
  const res = await fetch(url, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Outlook API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

export async function mailReply(
  messageId: string,
  options: { body: string; all?: boolean }
): Promise<void> {
  const action = options.all ? "createreplyall" : "createreply";
  // Create reply draft first
  const draft = (await outlookPost(`/messages/${messageId}/${action}`, {})) as MailMessage;

  // Prepare body with newlines preserved, prepend to existing quoted body
  const replyBody = prepareBody(options.body, false);
  const existingBody = draft.Body?.Content ?? "";
  const newContent = `${replyBody.Content}<br><br>${existingBody}`;

  // Update the draft body
  const data = (await outlookPatch(`/messages/${draft.Id}`, {
    Body: { ContentType: "HTML", Content: newContent },
  })) as MailMessage;

  console.log(`${c.green}Reply draft created.${c.reset}`);
  console.log(`${c.bold}Subject:${c.reset} ${data.Subject}`);
  const to = (data.ToRecipients ?? [])
    .map((r) => r.EmailAddress?.Address)
    .join(", ");
  console.log(`${c.bold}To:${c.reset}      ${to}`);
  const sid = shortId(data.Id);
  registerIds([data.Id]);
  console.log(`${c.bold}id:${c.reset}      ${c.yellow}${sid}${c.reset}`);
  console.log(`\nPreview: ${c.cyan}ms-cli mail open ${sid}${c.reset}`);
  console.log(`Send:    ${c.cyan}ms-cli mail send ${sid}${c.reset}`);
}

function prepareBody(text: string, html?: boolean): { ContentType: string; Content: string } {
  // Unescape literal \n from CLI args
  const unescaped = text.replace(/\\n/g, "\n");
  if (html) {
    return { ContentType: "HTML", Content: unescaped };
  }
  // Convert plain text to HTML to preserve newlines
  const escaped = unescaped
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");
  return { ContentType: "HTML", Content: `<div style="font-family:sans-serif">${escaped}</div>` };
}

function parseRecipient(addr: string): { EmailAddress: { Address: string; Name?: string } } {
  // "Name <email>" or just "email"
  const match = addr.match(/^(.+?)\s*<([^>]+)>$/);
  if (match) return { EmailAddress: { Name: match[1].trim(), Address: match[2] } };
  return { EmailAddress: { Address: addr } };
}

// --- Calendar ---

/** Fetch another user's calendar view, falling back to subscribed ICS calendar if user is external */
async function getUserCalendarView(
  email: string,
  start: string,
  end: string,
  top: number
): Promise<ODataResponse<CalendarEvent>> {
  const select = "subject,start,end,location,organizer,isAllDay,showAs,responseStatus,isCancelled";
  // Try Graph /users/{email} first (works for same-tenant users)
  try {
    const gdata = (await graphGet(
      `/users/${encodeURIComponent(email)}/calendarView?startDateTime=${start}&endDateTime=${end}&$top=${top}&$select=${select}&$orderby=start/dateTime`
    )) as { value: Array<Record<string, unknown>> };
    return { value: gdata.value.map(normalizeCalendarEvent) };
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    if (!msg.includes("ErrorInvalidUser") && !msg.includes("404")) throw e;
  }

  // Fallback: find subscribed ICS calendar
  const cal = await findSubscribedCalendar(email);
  if (!cal) {
    throw new Error(`ユーザー '${email}' が見つかりません（外部テナント、かつ購読カレンダーもありません）`);
  }
  console.error(`${c.dim}${email}: 購読カレンダー「${cal.name}」から取得${c.reset}`);
  const gdata = (await graphGet(
    `/me/calendars/${cal.id}/calendarView?startDateTime=${start}&endDateTime=${end}&$top=${top}&$select=${select}&$orderby=start/dateTime`
  )) as { value: Array<Record<string, unknown>> };
  return { value: gdata.value.map(normalizeCalendarEvent) };
}

/** Normalize Graph API camelCase response to Outlook PascalCase CalendarEvent */
function normalizeCalendarEvent(ev: Record<string, unknown>): CalendarEvent {
  const start = ev.start as { dateTime: string; timeZone: string } | undefined;
  const end = ev.end as { dateTime: string; timeZone: string } | undefined;
  const loc = ev.location as { displayName?: string } | undefined;
  const org = ev.organizer as { emailAddress?: { name: string; address: string } } | undefined;
  const resp = ev.responseStatus as { response: string } | undefined;
  return {
    Id: (ev.id as string) ?? "",
    Subject: (ev.subject as string) ?? "",
    Start: { DateTime: start?.dateTime ?? "", TimeZone: start?.timeZone ?? "" },
    End: { DateTime: end?.dateTime ?? "", TimeZone: end?.timeZone ?? "" },
    Location: loc ? { DisplayName: loc.displayName } : undefined,
    Organizer: org?.emailAddress ? { EmailAddress: { Name: org.emailAddress.name, Address: org.emailAddress.address } } : undefined,
    IsAllDay: (ev.isAllDay as boolean) ?? false,
    ShowAs: ev.showAs as string | undefined,
    ResponseStatus: resp ? { Response: resp.response } : undefined,
    IsCancelled: (ev.isCancelled as boolean) ?? false,
  };
}

interface CalendarEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  Location?: { DisplayName?: string };
  Organizer?: { EmailAddress: { Name: string; Address: string } };
  IsAllDay?: boolean;
  ShowAs?: string;
  ResponseStatus?: { Response: string };
  Body?: { ContentType: string; Content: string };
  Attendees?: Array<{
    EmailAddress: { Name: string; Address: string };
    Type: string;
    Status: { Response: string };
  }>;
  WebLink?: string;
  IsCancelled?: boolean;
}

export async function calendarList(options: {
  days?: number;
  pageSize?: number;
  user?: string;
}): Promise<void> {
  const days = options.days ?? 7;
  const pageSize = options.pageSize ?? 30;
  const now = new Date();
  const start = now.toISOString();
  const end = new Date(now.getTime() + days * 24 * 60 * 60 * 1000).toISOString();

  let data: ODataResponse<CalendarEvent>;
  if (options.user) {
    data = await getUserCalendarView(options.user, start, end, pageSize);
  } else {
    data = (await outlookGet(
      `/calendarview?startDateTime=${start}&endDateTime=${end}&$top=${pageSize}&$select=Subject,Start,End,Location,Organizer,IsAllDay,ShowAs,ResponseStatus,IsCancelled&$orderby=Start/DateTime`
    )) as ODataResponse<CalendarEvent>;
  }

  printAccountHeader();
  let currentDate = "";
  for (const ev of data.value) {
    if (ev.IsCancelled) continue;
    const s = new Date(ev.Start.DateTime + "Z");
    const e = new Date(ev.End.DateTime + "Z");
    const dateStr = s.toLocaleDateString("ja-JP", { timeZone: "Asia/Tokyo", month: "2-digit", day: "2-digit", weekday: "short" });

    if (dateStr !== currentDate) {
      currentDate = dateStr;
      console.log(`\n${c.bold}${c.cyan}── ${dateStr} ──${c.reset}`);
    }

    const showAs = ev.ShowAs ?? "Unknown";
    const showColor = showAs === "Busy" ? c.red : showAs === "Tentative" ? c.yellow : c.dim;
    const statusTag = `${showColor}[${showAs}]${c.reset}`;

    const timeStr = ev.IsAllDay
      ? `${c.magenta}終日${c.reset}          `
      : `${c.white}${s.toLocaleTimeString("ja-JP", { timeZone: "Asia/Tokyo", hour: "2-digit", minute: "2-digit" })} - ${e.toLocaleTimeString("ja-JP", { timeZone: "Asia/Tokyo", hour: "2-digit", minute: "2-digit" })}${c.reset}`;

    const loc = ev.Location?.DisplayName ? ` ${c.dim}@ ${ev.Location.DisplayName}${c.reset}` : "";
    const org = ev.Organizer?.EmailAddress?.Name ? ` ${c.green}(${ev.Organizer.EmailAddress.Name})${c.reset}` : "";
    const cancelled = ev.Subject?.startsWith("Canceled:") ? `${c.dim}${c.red}` : `${c.bold}`;

    console.log(`  ${statusTag} ${timeStr}  ${cancelled}${ev.Subject}${c.reset}${loc}${org}`);
  }

  if (options.user) {
    console.log(`\n${c.dim}(${options.user} の予定)${c.reset}`);
  }
  console.log(`${c.bold}${data.value.length}${c.reset} events (next ${days} days)`);
}

export async function calendarRead(
  eventId: string,
  options: { json?: boolean }
): Promise<void> {
  const data = (await outlookGet(
    `/events/${eventId}?$select=Subject,Start,End,Location,Organizer,Body,Attendees,IsAllDay,ShowAs,WebLink`
  )) as CalendarEvent;

  if (options.json) {
    console.log(JSON.stringify(data, null, 2));
    return;
  }

  const s = new Date(data.Start.DateTime + "Z");
  const e = new Date(data.End.DateTime + "Z");

  console.log(`${c.bold}Subject:${c.reset}   ${data.Subject}`);
  if (data.IsAllDay) {
    console.log(`${c.bold}Time:${c.reset}      ${c.magenta}終日${c.reset} ${s.toLocaleDateString("ja-JP", { timeZone: "Asia/Tokyo" })}`);
  } else {
    console.log(`${c.bold}Time:${c.reset}      ${s.toLocaleString("ja-JP", { timeZone: "Asia/Tokyo" })} - ${e.toLocaleTimeString("ja-JP", { timeZone: "Asia/Tokyo", hour: "2-digit", minute: "2-digit" })}`);
  }
  if (data.Location?.DisplayName) {
    console.log(`${c.bold}Location:${c.reset}  ${data.Location.DisplayName}`);
  }
  if (data.Organizer?.EmailAddress) {
    console.log(`${c.bold}Organizer:${c.reset} ${c.green}${data.Organizer.EmailAddress.Name} <${data.Organizer.EmailAddress.Address}>${c.reset}`);
  }
  if (data.ShowAs) {
    console.log(`${c.bold}Status:${c.reset}    ${data.ShowAs}`);
  }

  if (data.Attendees && data.Attendees.length > 0) {
    console.log(`${c.bold}Attendees:${c.reset}`);
    for (const a of data.Attendees) {
      const resp = a.Status?.Response ?? "None";
      const respColor = resp === "Accepted" ? c.green : resp === "Declined" ? c.red : c.yellow;
      const type = a.Type === "Required" ? "" : ` ${c.dim}(${a.Type})${c.reset}`;
      console.log(`  ${respColor}[${resp}]${c.reset} ${a.EmailAddress.Name ?? a.EmailAddress.Address}${type}`);
    }
  }

  if (data.WebLink) {
    console.log(`${c.bold}Link:${c.reset}      ${c.dim}${data.WebLink}${c.reset}`);
  }

  console.log(`${c.dim}${"─".repeat(60)}${c.reset}`);

  if (data.Body) {
    const content =
      data.Body.ContentType === "HTML"
        ? stripHtml(data.Body.Content)
        : data.Body.Content;
    console.log(content);
  }
}

export async function calendarToday(options?: { user?: string }): Promise<void> {
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const endOfDay = new Date(startOfDay.getTime() + 24 * 60 * 60 * 1000);
  const start = startOfDay.toISOString();
  const end = endOfDay.toISOString();

  let raw: ODataResponse<CalendarEvent>;
  if (options?.user) {
    raw = await getUserCalendarView(options.user, start, end, 50);
  } else {
    raw = (await outlookGet(
      `/calendarview?startDateTime=${start}&endDateTime=${end}&$top=50&$select=Subject,Start,End,Location,Organizer,IsAllDay,ShowAs,IsCancelled&$orderby=Start/DateTime`
    )) as ODataResponse<CalendarEvent>;
  }
  const data = { value: raw.value.filter((ev) => !ev.IsCancelled) };

  printAccountHeader();
  const userLabel = options?.user ? ` [${options.user}]` : "";
  console.log(`${c.bold}${c.cyan}── 今日の予定 (${now.toLocaleDateString("ja-JP", { month: "2-digit", day: "2-digit", weekday: "short" })})${userLabel} ──${c.reset}`);

  if (data.value.length === 0) {
    console.log(`  ${c.dim}予定なし${c.reset}`);
    return;
  }

  for (const ev of data.value) {
    const s = new Date(ev.Start.DateTime + "Z");
    const e = new Date(ev.End.DateTime + "Z");

    const showAs = ev.ShowAs ?? "Unknown";
    const showColor = showAs === "Busy" ? c.red : showAs === "Tentative" ? c.yellow : c.dim;
    const statusTag = `${showColor}[${showAs}]${c.reset}`;

    const timeStr = ev.IsAllDay
      ? `${c.magenta}終日${c.reset}          `
      : `${c.white}${s.toLocaleTimeString("ja-JP", { timeZone: "Asia/Tokyo", hour: "2-digit", minute: "2-digit" })} - ${e.toLocaleTimeString("ja-JP", { timeZone: "Asia/Tokyo", hour: "2-digit", minute: "2-digit" })}${c.reset}`;

    const loc = ev.Location?.DisplayName ? ` ${c.dim}@ ${ev.Location.DisplayName}${c.reset}` : "";
    const cancelled = ev.Subject?.startsWith("Canceled:") ? `${c.dim}${c.red}` : `${c.bold}`;

    console.log(`  ${statusTag} ${timeStr}  ${cancelled}${ev.Subject}${c.reset}${loc}`);
  }

  console.log(`\n${c.bold}${data.value.length}${c.reset} events`);
}

// --- Schedule / Find Slot ---

interface ScheduleItem {
  status: string;
  subject?: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
}

interface ScheduleInfo {
  scheduleId: string;
  availabilityView: string;
  scheduleItems: ScheduleItem[];
  workingHours?: {
    daysOfWeek: string[];
    startTime: string;
    endTime: string;
  };
  error?: {
    responseCode: string;
    message: string;
  };
}

/** Search user's subscribed calendars for one matching the email (ICS / shared) */
async function findSubscribedCalendar(
  email: string
): Promise<{ id: string; name: string } | null> {
  const localPart = email.split("@")[0].toLowerCase();
  try {
    const cals = (await graphGet(
      `/me/calendars?$top=100&$select=id,name,owner`
    )) as { value: Array<{ id: string; name: string; owner?: { address?: string; name?: string } }> };

    for (const cal of cals.value) {
      // Match by owner email
      if (cal.owner?.address?.toLowerCase() === email.toLowerCase()) {
        return { id: cal.id, name: cal.name };
      }
      // Match by calendar name containing the local part (e.g. "ANDO RYUICHI-TOKB042")
      const nameLower = cal.name.toLowerCase().replace(/[\s-]+/g, "");
      const searchPart = localPart.replace(/-/g, "");
      if (nameLower.includes(searchPart) && cal.name !== "Calendar") {
        return { id: cal.id, name: cal.name };
      }
    }
  } catch {}
  return null;
}

/** Build ScheduleInfo from calendarView events (fallback for ICS calendars) */
async function scheduleFromCalendarView(
  calendarId: string,
  calendarName: string,
  email: string,
  startDt: string,
  endDt: string,
  interval: number
): Promise<ScheduleInfo> {
  // Parse start/end as Asia/Tokyo
  const startMs = new Date(startDt + "+09:00").getTime();
  const endMs = new Date(endDt + "+09:00").getTime();
  const totalSlots = Math.floor((endMs - startMs) / (interval * 60 * 1000));

  const events = (await graphGet(
    `/me/calendars/${calendarId}/calendarView?startDateTime=${new Date(startMs).toISOString()}&endDateTime=${new Date(endMs).toISOString()}&$top=100&$select=subject,start,end,showAs,isAllDay,isCancelled&$orderby=start/dateTime`
  )) as { value: Array<{ subject?: string; start: { dateTime: string; timeZone: string }; end: { dateTime: string; timeZone: string }; showAs?: string; isCancelled?: boolean }> };

  // Build availability view (0=free, 1=tentative, 2=busy, 3=oof)
  const slots = Array.from({ length: totalSlots }, () => 0);
  const items: ScheduleItem[] = [];

  for (const ev of events.value) {
    if (ev.isCancelled) continue;
    const showAs = (ev.showAs ?? "busy").toLowerCase();
    if (showAs === "free") continue;

    const evStart = new Date(ev.start.dateTime + "Z").getTime();
    const evEnd = new Date(ev.end.dateTime + "Z").getTime();
    const statusCode = showAs === "tentative" ? 1 : showAs === "oof" ? 3 : 2;

    for (let i = 0; i < totalSlots; i++) {
      const slotStart = startMs + i * interval * 60 * 1000;
      const slotEnd = slotStart + interval * 60 * 1000;
      if (evStart < slotEnd && evEnd > slotStart) {
        slots[i] = Math.max(slots[i], statusCode);
      }
    }

    items.push({
      status: showAs,
      subject: ev.subject,
      start: ev.start,
      end: ev.end,
    });
  }

  return {
    scheduleId: email,
    availabilityView: slots.join(""),
    scheduleItems: items,
  };
}

// Cache: email → subscribed calendar (persists across calls within one process)
const icsCalendarCache = new Map<string, { id: string; name: string } | null>();

async function getSchedule(
  emails: string[],
  startDt: string,
  endDt: string,
  interval: number = 30
): Promise<ScheduleInfo[]> {
  // Split emails into ones we can query via getSchedule vs known ICS-only
  const graphEmails = emails.filter((e) => !icsCalendarCache.has(e));
  const icsEmails = emails.filter((e) => icsCalendarCache.has(e));

  // Query getSchedule for non-cached emails (skip if all are ICS)
  const results = new Map<string, ScheduleInfo>();
  if (graphEmails.length > 0) {
    const data = (await graphPost("/me/calendar/getSchedule", {
      schedules: graphEmails,
      startTime: { dateTime: startDt, timeZone: "Asia/Tokyo" },
      endTime: { dateTime: endDt, timeZone: "Asia/Tokyo" },
      availabilityViewInterval: interval,
    })) as { value: ScheduleInfo[] };

    for (const sched of data.value) {
      if (!sched.error) {
        results.set(sched.scheduleId, sched);
        continue;
      }

      // First-time error: try ICS fallback and cache result
      const cal = await findSubscribedCalendar(sched.scheduleId);
      icsCalendarCache.set(sched.scheduleId, cal);
      if (cal) {
        console.error(
          `${c.dim}${sched.scheduleId}: 購読カレンダー「${cal.name}」から取得${c.reset}`
        );
        try {
          results.set(sched.scheduleId, await scheduleFromCalendarView(
            cal.id, cal.name, sched.scheduleId, startDt, endDt, interval
          ));
          continue;
        } catch (e) {
          console.error(`${c.red}[Error]${c.reset} カレンダー「${cal.name}」の取得に失敗: ${e}`);
        }
      }
      // No fallback
      console.error(
        `${c.red}[Error]${c.reset} ${sched.scheduleId}: ${sched.error.message ?? sched.error.responseCode}`
      );
      results.set(sched.scheduleId, sched);
    }
  }

  // Fetch ICS-cached emails directly (no getSchedule call)
  for (const email of icsEmails) {
    const cal = icsCalendarCache.get(email);
    if (cal) {
      try {
        results.set(email, await scheduleFromCalendarView(
          cal.id, cal.name, email, startDt, endDt, interval
        ));
        continue;
      } catch (e) {
        console.error(`${c.red}[Error]${c.reset} カレンダー「${cal.name}」の取得に失敗: ${e}`);
      }
    }
    // Cached as null = no fallback available
    results.set(email, {
      scheduleId: email, availabilityView: "", scheduleItems: [],
      error: { responseCode: "NoCalendar", message: "購読カレンダーが見つかりません" },
    });
  }

  // Return in original email order
  return emails.map((e) => results.get(e)!);
}

export async function calendarSchedule(options: {
  emails: string[];
  date: string;
  startHour?: number;
  endHour?: number;
}): Promise<void> {
  const startHour = options.startHour ?? 9;
  const endHour = options.endHour ?? 17;
  const startDt = `${options.date}T${String(startHour).padStart(2, "0")}:00:00`;
  const endDt = `${options.date}T${String(endHour).padStart(2, "0")}:00:00`;

  const schedules = await getSchedule(options.emails, startDt, endDt);

  const totalSlots = (endHour - startHour) * 2; // 30min slots

  console.log(`${c.bold}${c.cyan}── ${options.date} (${startHour}:00-${endHour}:00) ──${c.reset}\n`);

  // Header: time slots
  const statusChars: Record<string, string> = {
    "0": `${c.green}░${c.reset}`,
    "1": `${c.yellow}▒${c.reset}`,
    "2": `${c.red}█${c.reset}`,
    "3": `${c.magenta}█${c.reset}`,
    "4": `${c.blue}▒${c.reset}`,
  };

  // Time ruler
  let ruler = "                          ";
  for (let h = startHour; h < endHour; h++) {
    ruler += `${String(h).padStart(2, "0")}  `;
  }
  console.log(`${c.dim}${ruler}${c.reset}`);

  for (const sched of schedules) {
    const name = sched.scheduleId.split("@")[0].padEnd(25);
    if (sched.error) {
      console.log(`${c.bold}${name}${c.reset} ${c.red}(エラー: ${sched.error.message ?? sched.error.responseCode})${c.reset}`);
      continue;
    }
    const view = (sched.availabilityView ?? "").slice(0, totalSlots);
    let bar = "";
    for (const ch of view) {
      bar += statusChars[ch] ?? ch;
    }
    console.log(`${c.bold}${name}${c.reset} ${bar}`);

    // Show busy items
    for (const item of (sched.scheduleItems ?? [])) {
      if (item.status === "free") continue;
      const s = new Date(item.start.dateTime + "Z");
      const e = new Date(item.end.dateTime + "Z");
      const sTime = s.toLocaleTimeString("ja-JP", { timeZone: "Asia/Tokyo", hour: "2-digit", minute: "2-digit" });
      const eTime = e.toLocaleTimeString("ja-JP", { timeZone: "Asia/Tokyo", hour: "2-digit", minute: "2-digit" });
      const statusColor = item.status === "busy" ? c.red : item.status === "oof" ? c.magenta : c.yellow;
      console.log(`${c.dim}                          ${statusColor}[${item.status}]${c.reset} ${sTime}-${eTime} ${item.subject ?? "(private)"}${c.reset}`);
    }
  }

  console.log(`\n${c.dim}Legend: ${c.green}░${c.reset}${c.dim}=Free ${c.yellow}▒${c.reset}${c.dim}=Tentative ${c.red}█${c.reset}${c.dim}=Busy ${c.magenta}█${c.reset}${c.dim}=OOF${c.reset}`);
}

export async function calendarFindSlot(options: {
  emails: string[];
  duration: number;
  days?: number;
  startHour?: number;
  endHour?: number;
  startDate?: string;
}): Promise<void> {
  const days = options.days ?? 5;
  const duration = options.duration;
  const startHour = options.startHour ?? 9;
  const endHour = options.endHour ?? 17;
  const slotsNeeded = Math.ceil(duration / 30);

  const today = options.startDate
    ? new Date(options.startDate + "T00:00:00+09:00")
    : new Date();

  const results: Array<{ date: string; start: string; end: string }> = [];

  // Check day by day (skip weekends)
  let checked = 0;
  let dayOffset = 0;
  while (checked < days) {
    const d = new Date(today.getTime() + dayOffset * 24 * 60 * 60 * 1000);
    dayOffset++;
    const dow = d.getDay();
    if (dow === 0 || dow === 6) continue; // skip weekends
    checked++;

    const dateStr = d.toLocaleDateString("sv-SE", { timeZone: "Asia/Tokyo" }); // YYYY-MM-DD
    const startDt = `${dateStr}T${String(startHour).padStart(2, "0")}:00:00`;
    const endDt = `${dateStr}T${String(endHour).padStart(2, "0")}:00:00`;

    const schedules = await getSchedule(options.emails, startDt, endDt);
    const totalSlots = (endHour - startHour) * 2;

    // Merge availability: a slot is free only if ALL users are free (0)
    const merged = Array.from({ length: totalSlots }, () => true);
    for (const sched of schedules) {
      if (sched.error || !sched.availabilityView) {
        // Treat errored user as fully busy (cannot confirm availability)
        merged.fill(false);
        continue;
      }
      for (let i = 0; i < totalSlots && i < sched.availabilityView.length; i++) {
        if (sched.availabilityView[i] !== "0") merged[i] = false;
      }
    }

    // Find consecutive free slots
    let consecutive = 0;
    for (let i = 0; i < totalSlots; i++) {
      if (merged[i]) {
        consecutive++;
        if (consecutive >= slotsNeeded) {
          const slotStart = i - slotsNeeded + 1;
          const startMin = startHour * 60 + slotStart * 30;
          const endMin = startMin + duration;
          const sH = String(Math.floor(startMin / 60)).padStart(2, "0");
          const sM = String(startMin % 60).padStart(2, "0");
          const eH = String(Math.floor(endMin / 60)).padStart(2, "0");
          const eM = String(endMin % 60).padStart(2, "0");
          results.push({ date: dateStr, start: `${sH}:${sM}`, end: `${eH}:${eM}` });
          consecutive = 0; // reset to find next non-overlapping slot
        }
      } else {
        consecutive = 0;
      }
    }
  }

  console.log(`${c.bold}${c.cyan}── 空きスロット検索 ──${c.reset}`);
  console.log(`${c.dim}参加者: ${options.emails.join(", ")}${c.reset}`);
  console.log(`${c.dim}時間帯: ${startHour}:00-${endHour}:00 / ${duration}分 / 平日${days}日間${c.reset}\n`);

  if (results.length === 0) {
    console.log(`${c.red}共通の空きスロットが見つかりませんでした${c.reset}`);
    return;
  }

  let currentDate = "";
  for (const slot of results) {
    if (slot.date !== currentDate) {
      currentDate = slot.date;
      const d = new Date(slot.date + "T00:00:00+09:00");
      const dayName = d.toLocaleDateString("ja-JP", { timeZone: "Asia/Tokyo", weekday: "short" });
      console.log(`${c.bold}${slot.date} (${dayName})${c.reset}`);
    }
    console.log(`  ${c.green}${slot.start} - ${slot.end}${c.reset}`);
  }

  console.log(`\n${c.bold}${results.length}${c.reset} slots found`);
}

// --- Mail Open (preview in browser) ---

export async function mailOpen(messageId: string): Promise<void> {
  const data = (await outlookGet(
    `/messages/${messageId}?$select=Subject,WebLink`
  )) as MailMessage & { WebLink?: string };

  if (!data.WebLink) {
    console.error("WebLink not available for this message.");
    process.exit(1);
  }

  console.log(`${c.bold}Opening:${c.reset} ${data.Subject}`);
  execSync(`open ${JSON.stringify(data.WebLink)}`);
}

// --- Helpers ---

function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/li>/gi, "\n")
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
    return (
      d.toLocaleDateString("ja-JP", { month: "2-digit", day: "2-digit" }) +
      " " +
      d.toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" })
    );
  } catch {
    return iso;
  }
}
