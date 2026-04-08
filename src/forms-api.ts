import { loadConfig, saveConfig, type Config } from "./config.js";

const TEAMS_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
const FORMS_BASE = "https://forms.office.com/formapi/api";

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
};

function decodeJwt(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3 && parts.length !== 5) return null;
    const payload = parts[1];
    return JSON.parse(Buffer.from(payload + "==", "base64url").toString());
  } catch {
    return null;
  }
}

/** Get oid (user object id) from any available JWT token in config */
function getUserId(config: { graphToken?: string; outlookToken?: string; skypeToken?: string }): string {
  for (const token of [config.graphToken, config.outlookToken, config.skypeToken]) {
    if (!token) continue;
    const payload = decodeJwt(token);
    if (payload?.oid) return payload.oid as string;
  }
  return "";
}

async function ensureFormsToken(): Promise<{ token: string; tenantId: string; userId: string }> {
  const config = loadConfig();
  const tenantId = config.tenantId ?? "common";

  // Forms tokens are opaque (encrypted JWE, 5 parts) — can't decode exp.
  // Track expiry separately via formsTokenExp in config.
  if (config.formsToken && config.formsTokenExp && config.formsTokenExp > Date.now() / 1000 + 60) {
    const userId = config.formsUserId ?? getUserId(config);
    return { token: config.formsToken, tenantId, userId };
  }

  if (!config.refreshToken) {
    console.error("No refresh token. Run: ms-cli auth login");
    process.exit(1);
  }

  console.error("Refreshing Forms token...");
  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Origin: "https://teams.microsoft.com",
      },
      body: new URLSearchParams({
        client_id: TEAMS_CLIENT_ID,
        grant_type: "refresh_token",
        refresh_token: config.refreshToken,
        scope: "https://forms.office.com/.default openid profile offline_access",
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
    id_token?: string;
    expires_in?: number;
  };

  config.formsToken = data.access_token;
  config.formsTokenExp = Math.floor(Date.now() / 1000) + (data.expires_in ?? 3600);

  // Extract oid from id_token (which IS a standard JWT)
  if (data.id_token) {
    const idPayload = decodeJwt(data.id_token);
    if (idPayload?.oid) config.formsUserId = idPayload.oid as string;
  }
  if (!config.formsUserId) {
    config.formsUserId = getUserId(config);
  }

  if (data.refresh_token) {
    config.refreshToken = data.refresh_token;
    config.refreshTokenIssuedAt = Math.floor(Date.now() / 1000);
  }
  saveConfig(config);
  console.error("Forms token refreshed.");

  return {
    token: data.access_token,
    tenantId,
    userId: config.formsUserId ?? "",
  };
}

async function formsGet(path: string): Promise<unknown> {
  const { token } = await ensureFormsToken();
  const url = path.startsWith("http") ? path : `${FORMS_BASE}${path}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Forms API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

// --- Types ---

interface SharedForm {
  id: string;
  title: string;
  ownerId: string;
  ownerTenantId: string;
  ownerDisplayName?: string;
  dateTimeShared?: string;
  defaultAppRestriction?: string;
}

interface FormDetail {
  id: string;
  title: string;
  formsProRTDescription?: string;
  createdDate?: string;
  modifiedDate?: string;
  softDeleted?: number;
  ownerId: string;
  ownerTenantId: string;
  settings?: string; // JSON string
  type?: string;
}

interface FormQuestion {
  id: string;
  title: string;
  type: string; // "Question.Choice", "Question.Text", "Question.Rating", etc.
  required?: boolean;
  order?: number;
  questionInfo?: string; // JSON string with choices etc.
}

interface FormResponse {
  id: number;
  startDate?: string;
  submitDate?: string;
  responder?: string;
  responderName?: string;
  answers?: string; // JSON string: [{answer1, questionId}, ...]
}

interface AnswerEntry {
  questionId: string;
  answer1?: string;
  answer2?: string;
}

/** Parse questionInfo JSON to extract choices */
function parseChoices(questionInfo?: string): string[] {
  if (!questionInfo) return [];
  try {
    const info = JSON.parse(questionInfo);
    if (info.Choices && Array.isArray(info.Choices)) {
      return info.Choices.map((ch: { Description?: string }) => ch.Description ?? "");
    }
  } catch {}
  return [];
}

/** Strip HTML tags */
function stripHtml(html: string): string {
  return html.replace(/<[^>]*>/g, "").trim();
}

// --- Parse form URL to extract formId and resolve ownerId ---

/** Extract form id from a Forms URL */
export function parseFormsUrl(url: string): string | null {
  try {
    const u = new URL(url);
    return u.searchParams.get("id");
  } catch {
    return null;
  }
}

/** Resolve the ownerId for a form by checking sharedWithMeForms */
async function resolveFormOwner(
  tenantId: string,
  userId: string,
  formId: string
): Promise<string | null> {
  try {
    const data = (await formsGet(
      `/${tenantId}/users/${userId}/sharedWithMeForms`
    )) as { value: SharedForm[] };

    const match = data.value?.find((f) => f.id === formId);
    return match?.ownerId ?? null;
  } catch {
    return null;
  }
}

// --- List Forms (own + shared) ---

export async function formsList(options: {
  pageSize?: number;
  json?: boolean;
}): Promise<void> {
  const { tenantId, userId } = await ensureFormsToken();

  // Fetch own forms and shared forms in parallel
  const [ownData, sharedData] = await Promise.all([
    formsGet(`/${tenantId}/users/${userId}/light/forms`) as Promise<{ value: Array<{ id: string; title: string; modifiedDate?: string; responseCount?: number }> }>,
    formsGet(`/${tenantId}/users/${userId}/sharedWithMeForms`) as Promise<{ value: SharedForm[] }>,
  ]);

  const ownForms = ownData.value ?? [];
  const sharedForms = sharedData.value ?? [];

  if (options.json) {
    console.log(JSON.stringify({ own: ownForms, shared: sharedForms }, null, 2));
    return;
  }

  const pageSize = options.pageSize ?? 20;

  if (ownForms.length > 0) {
    console.log(`${c.bold}${c.cyan}── 自分のフォーム ──${c.reset}\n`);
    for (const form of ownForms.slice(0, pageSize)) {
      const responses = form.responseCount != null ? ` ${c.cyan}${form.responseCount} responses${c.reset}` : "";
      const modified = form.modifiedDate ? formatTime(form.modifiedDate) : "";
      console.log(`  ${c.bold}${form.title || "(無題)"}${c.reset}${responses}`);
      console.log(`  ${c.dim}id: ${form.id}${c.reset}`);
      if (modified) console.log(`  ${c.dim}updated: ${modified}${c.reset}`);
      console.log();
    }
  }

  if (sharedForms.length > 0) {
    console.log(`${c.bold}${c.cyan}── 共有されたフォーム ──${c.reset}\n`);
    for (const form of sharedForms.slice(0, pageSize)) {
      const owner = form.ownerDisplayName ?? form.ownerId;
      const shared = form.dateTimeShared ? formatTime(form.dateTimeShared) : "";
      console.log(`  ${c.bold}${form.title || "(無題)"}${c.reset} ${c.dim}by ${owner}${c.reset}`);
      console.log(`  ${c.dim}id: ${form.id}${c.reset}`);
      console.log(`  ${c.dim}owner: ${form.ownerId}${c.reset}`);
      if (shared) console.log(`  ${c.dim}shared: ${shared}${c.reset}`);
      console.log();
    }
  }

  if (ownForms.length === 0 && sharedForms.length === 0) {
    console.log("No forms found.");
    return;
  }

  console.log(`${c.bold}${ownForms.length}${c.reset} own, ${c.bold}${sharedForms.length}${c.reset} shared`);
}

// --- Form Detail ---

export async function formsRead(
  formIdOrUrl: string,
  options: { owner?: string; json?: boolean }
): Promise<void> {
  const { tenantId, userId } = await ensureFormsToken();

  // If it looks like a URL, parse out the form ID
  const formId = formIdOrUrl.startsWith("http") ? (parseFormsUrl(formIdOrUrl) ?? formIdOrUrl) : formIdOrUrl;

  // Determine owner: explicit flag > resolve from sharedWithMe > self
  let ownerId = options.owner ?? "";
  if (!ownerId) {
    ownerId = (await resolveFormOwner(tenantId, userId, formId)) ?? userId;
  }

  const data = (await formsGet(
    `/${tenantId}/users/${ownerId}/forms('${formId}')`
  )) as FormDetail;

  // Get questions separately
  const qData = (await formsGet(
    `/${tenantId}/users/${ownerId}/forms('${formId}')/questions`
  )) as { value: FormQuestion[] };
  const questions = qData.value ?? [];

  if (options.json) {
    console.log(JSON.stringify({ ...data, questions }, null, 2));
    return;
  }

  const title = data.formsProRTDescription
    ? stripHtml(data.formsProRTDescription)
    : "";

  console.log(`${c.bold}${c.cyan}── ${stripHtml(data.title)} ──${c.reset}`);
  if (title) console.log(`${c.dim}${title}${c.reset}`);
  console.log(`${c.bold}id:${c.reset}      ${c.dim}${data.id}${c.reset}`);
  console.log(`${c.bold}owner:${c.reset}   ${c.dim}${data.ownerId}${c.reset}`);
  if (data.modifiedDate) {
    console.log(`${c.bold}updated:${c.reset} ${formatTime(data.modifiedDate)}`);
  }

  if (questions.length > 0) {
    console.log(`\n${c.bold}Questions:${c.reset}`);
    const sorted = [...questions].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
    for (let i = 0; i < sorted.length; i++) {
      const q = sorted[i];
      const req = q.required ? ` ${c.red}*${c.reset}` : "";
      const typeLabel = q.type.replace("Question.", "");
      console.log(`  ${c.yellow}Q${i + 1}.${c.reset} ${c.bold}${q.title}${c.reset}${req} ${c.dim}[${typeLabel}]${c.reset}`);

      const choices = parseChoices(q.questionInfo);
      for (const ch of choices) {
        console.log(`      ${c.dim}○ ${ch}${c.reset}`);
      }
    }
  }
}

// --- Form Responses ---

export async function formsResponses(
  formIdOrUrl: string,
  options: { owner?: string; pageSize?: number; json?: boolean }
): Promise<void> {
  const { tenantId, userId } = await ensureFormsToken();

  const formId = formIdOrUrl.startsWith("http") ? (parseFormsUrl(formIdOrUrl) ?? formIdOrUrl) : formIdOrUrl;

  let ownerId = options.owner ?? "";
  if (!ownerId) {
    ownerId = (await resolveFormOwner(tenantId, userId, formId)) ?? userId;
  }

  // Fetch responses and questions in parallel
  const [respData, qData] = await Promise.all([
    formsGet(`/${tenantId}/users/${ownerId}/forms('${formId}')/responses`) as Promise<{ value: FormResponse[] }>,
    formsGet(`/${tenantId}/users/${ownerId}/forms('${formId}')/questions`) as Promise<{ value: FormQuestion[] }>,
  ]);

  const responses = respData.value ?? [];
  const questions = (qData.value ?? []).sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

  if (options.json) {
    console.log(JSON.stringify({ questions, responses }, null, 2));
    return;
  }

  if (responses.length === 0) {
    console.log("No responses.");
    return;
  }

  // Build questionId -> title map
  const qMap = new Map(questions.map((q) => [q.id, q.title]));

  const pageSize = options.pageSize ?? 20;
  const shown = responses.slice(0, pageSize);

  for (let i = 0; i < shown.length; i++) {
    const resp = shown[i];
    const responder = resp.responderName || (resp.responder === "anonymous" ? "Anonymous" : resp.responder ?? "Unknown");
    const time = resp.submitDate ? formatTime(resp.submitDate) : "";

    console.log(`${c.bold}${c.cyan}── #${resp.id} ──${c.reset} ${c.green}${responder}${c.reset} ${c.dim}(${time})${c.reset}`);

    if (resp.answers) {
      try {
        const answers: AnswerEntry[] = typeof resp.answers === "string" ? JSON.parse(resp.answers) : resp.answers;
        if (Array.isArray(answers)) {
          for (const a of answers) {
            const label = qMap.get(a.questionId) ?? a.questionId;
            const value = a.answer1 ?? "";
            console.log(`  ${c.yellow}${label}:${c.reset} ${value}`);
          }
        }
      } catch {
        console.log(`  ${c.dim}${String(resp.answers).slice(0, 200)}${c.reset}`);
      }
    }
    console.log();
  }

  console.log(`${c.bold}${shown.length}${c.reset}${responses.length > pageSize ? ` / ${responses.length}` : ""} responses`);
}

// --- Group Forms ---

export async function formsGroup(
  groupId: string,
  options: { pageSize?: number; json?: boolean }
): Promise<void> {
  const data = (await formsGet(
    `/groups/${groupId}/forms`
  )) as { value: Array<{ id: string; title: string; responseCount?: number }> };

  const forms = data.value ?? [];

  if (options.json) {
    console.log(JSON.stringify(forms, null, 2));
    return;
  }

  if (forms.length === 0) {
    console.log("No group forms found.");
    return;
  }

  const pageSize = options.pageSize ?? 20;
  const shown = forms.slice(0, pageSize);

  for (const form of shown) {
    const responses = form.responseCount != null ? ` ${c.cyan}${form.responseCount} responses${c.reset}` : "";
    console.log(`  ${c.bold}${form.title || "(無題)"}${c.reset}${responses}`);
    console.log(`  ${c.dim}id: ${form.id}${c.reset}`);
    console.log();
  }

  console.log(`${c.bold}${shown.length}${c.reset} group forms`);
}

// --- Helpers ---

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
