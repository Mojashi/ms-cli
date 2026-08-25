import {
  loadConfig,
  saveConfig,
  getConfigPath,
  saveAccount,
  listAccounts,
  useAccount,
  removeAccount,
  getCurrentAccountKey,
  type Config,
} from "./config.js";
import { acquireInteractiveToken } from "./oauth-webview.js";

interface JwtPayload {
  iat?: number;
  exp?: number;
  skypeid?: string;
  scp?: number;
  rgn?: string;
  tid?: string;
  aud?: string;
  upn?: string;
  unique_name?: string;
  [key: string]: unknown;
}

/** Best-effort email/UPN for an account, decoded from whatever token is present. */
function accountUpn(config: Config): string | undefined {
  for (const tok of [config.outlookToken, config.graphToken, config.skypeToken]) {
    if (!tok) continue;
    try {
      const p = decodeJwt(tok);
      if (p.upn) return p.upn;
      if (p.unique_name) return p.unique_name;
    } catch {}
  }
  return undefined;
}

function decodeJwt(token: string): JwtPayload {
  const parts = token.split(".");
  if (parts.length !== 3) throw new Error("Invalid JWT");
  const payload = parts[1];
  const padded = payload + "=".repeat((4 - (payload.length % 4)) % 4);
  return JSON.parse(Buffer.from(padded, "base64url").toString());
}

export function isTokenValid(): boolean {
  const config = loadConfig();
  if (!config.skypeToken) return false;
  try {
    const payload = decodeJwt(config.skypeToken);
    const now = Math.floor(Date.now() / 1000);
    return (payload.exp ?? 0) > now;
  } catch {
    return false;
  }
}

export function tokenStatus(): void {
  const config = loadConfig();
  if (!config.skypeToken) {
    console.log("Not logged in. Run: ms-cli auth login");
    return;
  }
  try {
    const payload = decodeJwt(config.skypeToken);
    const now = Math.floor(Date.now() / 1000);
    const exp = payload.exp ?? 0;
    const remaining = exp - now;

    const acctKey = getCurrentAccountKey();
    if (acctKey) console.log(`account:  ${acctKey}`);
    console.log(`skypeid:  ${payload.skypeid ?? "unknown"}`);
    console.log(`region:   ${payload.rgn ?? "unknown"}`);
    console.log(`tenant:   ${payload.tid ?? "unknown"}`);
    console.log(`expires:  ${new Date(exp * 1000).toISOString()}`);
    if (remaining > 0) {
      const h = Math.floor(remaining / 3600);
      const m = Math.floor((remaining % 3600) / 60);
      console.log(`remaining: ${h}h ${m}m`);
    } else {
      console.log(`status:   EXPIRED (${Math.floor(-remaining / 60)}m ago)`);
    }
    // Refresh token info
    if (config.refreshToken) {
      if (config.refreshTokenIssuedAt) {
        const RT_LIFETIME = 90 * 24 * 3600; // 90 days
        const rtExp = config.refreshTokenIssuedAt + RT_LIFETIME;
        const rtRemaining = rtExp - now;
        const rtExpDate = new Date(rtExp * 1000).toISOString();
        if (rtRemaining > 0) {
          const d = Math.floor(rtRemaining / 86400);
          const h = Math.floor((rtRemaining % 86400) / 3600);
          console.log(`refresh:  expires ~${rtExpDate} (~${d}d ${h}h remaining)`);
        } else {
          console.log(`refresh:  EXPIRED (~${Math.floor(-rtRemaining / 86400)}d ago)`);
        }
      } else {
        console.log(`refresh:  present (issued date unknown)`);
      }
    } else {
      console.log(`refresh:  none`);
    }
    console.log(`config:   ${getConfigPath()}`);
  } catch (e) {
    console.error("Failed to decode token:", (e as Error).message);
  }
}

export function login(skypeToken: string, refreshToken?: string): void {
  const config = loadConfig();
  config.skypeToken = skypeToken.trim();
  if (refreshToken) {
    config.refreshToken = refreshToken.trim();
    config.refreshTokenIssuedAt = Math.floor(Date.now() / 1000);
  }

  // auto-detect region from token
  try {
    const payload = decodeJwt(config.skypeToken);
    if (payload.rgn) {
      config.region = payload.rgn;
      config.chatServiceHost = `${payload.rgn}.ng.msg.teams.microsoft.com`;
    }
    if (payload.tid) config.tenantId = payload.tid;
  } catch {}

  saveConfig(config);
  console.log("Token saved.");
  tokenStatus();
}

const TEAMS_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"; // Microsoft Teams native client

const SPACES_SCOPE =
  "https://api.spaces.skype.com/.default openid profile offline_access";
const AUTHZ_ENDPOINTS = [
  "https://teams.microsoft.com/api/authsvc/v1.0/authz",
  "https://authsvc.teams.microsoft.com/v1.0/authz",
];

/** Get an api.spaces.skype.com AAD token from a config's refresh token. */
async function getSpacesToken(config: Config): Promise<string | null> {
  if (!config.refreshToken) return null;
  const tenantId = config.tenantId ?? "common";
  const clientId = config.clientId ?? TEAMS_CLIENT_ID;
  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        grant_type: "refresh_token",
        refresh_token: config.refreshToken,
        scope: SPACES_SCOPE,
      }),
    }
  );
  if (!res.ok) return null;
  return ((await res.json()) as { access_token?: string }).access_token ?? null;
}

export interface TenantInfo {
  tenantId: string;
  tenantName: string;
  userType: string;
}

/** List all tenants (home + guest) the user can access, via the Teams MT API. */
async function fetchTenantList(spacesToken: string): Promise<TenantInfo[]> {
  for (const region of ["apac", "emea", "noam", "amer"]) {
    try {
      const r = await fetch(
        `https://teams.microsoft.com/api/mt/${region}/beta/users/tenantsv2`,
        { headers: { Authorization: `Bearer ${spacesToken}` } }
      );
      if (r.ok) {
        const arr = (await r.json()) as Array<{
          tenantId: string;
          tenantName: string;
          userType: string;
        }>;
        return arr.map((t) => ({
          tenantId: t.tenantId,
          tenantName: t.tenantName,
          userType: t.userType,
        }));
      }
    } catch {}
  }
  return [];
}

/** Mint a self-contained account config for a tenant, reusing a base refresh token. */
async function buildTenantAccount(
  base: Config,
  tenantId: string
): Promise<Config | null> {
  if (!base.refreshToken) return null;
  const clientId = base.clientId ?? TEAMS_CLIENT_ID;
  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        grant_type: "refresh_token",
        refresh_token: base.refreshToken,
        scope: SPACES_SCOPE,
      }),
    }
  );
  if (!res.ok) return null;
  const data = (await res.json()) as { access_token: string; refresh_token?: string };
  const cfg: Config = {
    skypeToken: "",
    clientId,
    tenantId,
    refreshToken: data.refresh_token ?? base.refreshToken,
    refreshTokenIssuedAt: Math.floor(Date.now() / 1000),
  };
  for (const ep of AUTHZ_ENDPOINTS) {
    try {
      const r = await fetch(ep, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${data.access_token}`,
          "Content-Type": "application/json",
        },
        body: "{}",
      });
      if (r.ok) {
        const j = (await r.json()) as { tokens?: { skypeToken?: string } };
        if (j.tokens?.skypeToken) {
          cfg.skypeToken = j.tokens.skypeToken;
          try {
            const p = decodeJwt(cfg.skypeToken);
            if (p.rgn) {
              cfg.region = p.rgn;
              cfg.chatServiceHost = `${p.rgn}.ng.msg.teams.microsoft.com`;
            }
          } catch {}
          break;
        }
      }
    } catch {}
  }
  return cfg;
}

/** Discover all tenants for the current account and register any missing ones. */
export async function syncTenants(): Promise<void> {
  const base = loadConfig();
  if (!base.refreshToken) {
    console.error("Not logged in. Run: ms-cli auth login");
    process.exit(1);
  }
  const spaces = await getSpacesToken(base);
  if (!spaces) {
    console.error("Could not obtain a token for tenant discovery.");
    process.exit(1);
  }
  const tenants = await fetchTenantList(spaces);
  if (tenants.length === 0) {
    console.error("No tenants discovered (API returned none).");
    return;
  }
  console.log(`Discovered ${tenants.length} tenant(s):`);
  for (const t of tenants) {
    const existing = listAccounts().find((a) => a.config.tenantId === t.tenantId);
    if (existing) {
      // Metadata-only accounts can remain after migration or an interrupted
      // login. Rebuild their tenant-specific refresh token instead of treating
      // them as fully registered.
      if (!existing.config.refreshToken) {
        process.stdout.write(
          `  ~ ${t.tenantName} [${t.userType}] ... recovering authentication ... `
        );
        const recovered = await buildTenantAccount(base, t.tenantId);
        if (!recovered) {
          console.log("failed (token exchange)");
          continue;
        }
        saveAccount(
          existing.key,
          {
            ...existing.config,
            ...recovered,
            tenantName: t.tenantName,
            userType: t.userType,
          },
          false
        );
        console.log(
          recovered.skypeToken
            ? "recovered"
            : "recovered (no Teams token)"
        );
        continue;
      }

      // Backfill the friendly tenant name / user type onto existing accounts.
      if (existing.config.tenantName !== t.tenantName || existing.config.userType !== t.userType) {
        saveAccount(existing.key, { ...existing.config, tenantName: t.tenantName, userType: t.userType }, false);
      }
      console.log(`  = ${t.tenantName} [${t.userType}] — already registered as "${existing.key}"`);
      continue;
    }
    process.stdout.write(`  + ${t.tenantName} [${t.userType}] ... `);
    const cfg = await buildTenantAccount(base, t.tenantId);
    if (!cfg) {
      console.log("failed (token exchange)");
      continue;
    }
    cfg.tenantName = t.tenantName;
    cfg.userType = t.userType;
    saveAccount(t.tenantName, cfg, false);
    console.log(cfg.skypeToken ? "registered" : "registered (no Teams token)");
  }
  console.log();
  authList();
}

/** Try to refresh skypetoken using saved refresh token. Returns true on success. */
export async function tryRefresh(quiet = false): Promise<boolean> {
  const config = loadConfig();
  if (!config.refreshToken) {
    if (!quiet) console.error("No refresh token saved. Run: ms-cli auth login");
    return false;
  }
  const clientId = config.clientId ?? TEAMS_CLIENT_ID;

  const tenantId = config.tenantId ?? "common";

  // Step 1: Refresh AAD token for api.spaces.skype.com
  if (!quiet) console.log("Refreshing AAD token...");
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "refresh_token",
    refresh_token: config.refreshToken,
    scope: "https://api.spaces.skype.com/.default openid profile offline_access",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    if (!quiet) {
      const text = await res.text();
      console.error(`AAD token refresh failed (${res.status}): ${text}`);
    }
    return false;
  }

  const aadData = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
  };
  if (!quiet) console.log("AAD token refreshed.");

  // Save new refresh token if rotated
  if (aadData.refresh_token) {
    config.refreshToken = aadData.refresh_token;
    config.refreshTokenIssuedAt = Math.floor(Date.now() / 1000);
  }

  // Step 2: Exchange AAD token for skypetoken
  if (!quiet) console.log("Exchanging for skypetoken...");

  const authzEndpoints = [
    "https://teams.microsoft.com/api/authsvc/v1.0/authz",
    "https://authsvc.teams.microsoft.com/v1.0/authz",
  ];

  for (const endpoint of authzEndpoints) {
    try {
      const skypeRes = await fetch(endpoint, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${aadData.access_token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({}),
      });

      if (skypeRes.ok) {
        const skypeData = (await skypeRes.json()) as { tokens?: { skypeToken?: string } };
        if (skypeData.tokens?.skypeToken) {
          config.skypeToken = skypeData.tokens.skypeToken;
          saveConfig(config);
          if (!quiet) {
            console.log("Skypetoken refreshed.");
            tokenStatus();
          }
          return true;
        }
      }
    } catch {}
  }

  // If skypetoken exchange fails, at least save the refreshed refresh token
  saveConfig(config);
  if (!quiet) {
    console.error(
      "Could not exchange AAD token for skypetoken. AAD refresh token was updated." +
      "\nYou may need to manually paste a new skypetoken."
    );
  }
  return false;
}

export async function refresh(): Promise<void> {
  const success = await tryRefresh(false);
  if (!success) process.exit(1);
}

/**
 * Authorization-code + PKCE login in a macOS WebView. Logs into a possibly
 * new account without clobbering existing accounts. Tokens are stored in the
 * regular multi-account config after the interactive login succeeds.
 */
export async function webViewLogin(
  opts: { name?: string; tenant?: string } = {}
): Promise<void> {
  // Start from a fresh account so a second login never overwrites the current one.
  const config: Config = { skypeToken: "" };
  let upn: string | undefined;
  const clientId = TEAMS_CLIENT_ID;
  const tenantId = opts.tenant ?? "common";

  console.log("Opening Microsoft 365 login...");
  const tokenData = await acquireInteractiveToken(
    clientId,
    tenantId,
    SPACES_SCOPE
  );
  config.refreshToken = tokenData.refreshToken;
  config.refreshTokenIssuedAt = Math.floor(Date.now() / 1000);

  // Exchange the AAD access token for a Teams skypetoken.
  try {
    const payload = decodeJwt(tokenData.accessToken);
    upn = payload.upn ?? payload.unique_name;
    if (payload.tid) config.tenantId = payload.tid;
  } catch {}
  console.log("Got AAD token. Exchanging for skypetoken...");
  for (const endpoint of AUTHZ_ENDPOINTS) {
    try {
      const skypeResponse = await fetch(endpoint, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${tokenData.accessToken}`,
          "Content-Type": "application/json",
        },
        body: "{}",
      });
      if (!skypeResponse.ok) continue;
      const skypeData = (await skypeResponse.json()) as {
        tokens?: { skypeToken?: string };
      };
      if (skypeData.tokens?.skypeToken) {
        config.skypeToken = skypeData.tokens.skypeToken;
        break;
      }
    } catch {}
  }

  // Detect region/tenant from the freshly-issued skypetoken.
  try {
    const payload = decodeJwt(config.skypeToken);
    if (payload.rgn) {
      config.region = payload.rgn;
      config.chatServiceHost = `${payload.rgn}.ng.msg.teams.microsoft.com`;
    }
    if (payload.tid) config.tenantId = payload.tid;
  } catch {}

  // Account key. Default to the UPN, but the SAME UPN can exist in multiple
  // tenants (guest/B2B), so disambiguate by tenant to avoid clobbering a
  // different tenant's account that happens to share the email.
  const baseKey = opts.name ?? upn ?? config.tenantId ?? "default";
  let key = baseKey;
  if (!opts.name) {
    const clash = listAccounts().some(
      (account) =>
        account.key === baseKey &&
        account.config.tenantId &&
        config.tenantId &&
        account.config.tenantId !== config.tenantId
    );
    if (clash) key = `${baseKey} (${(config.tenantId ?? "").slice(0, 8)})`;
  }

  // Re-login keeps still-valid derived resource tokens. Fresh values win, but
  // an authz failure must not clobber an existing working skypetoken.
  const existing = listAccounts().find(
    (account) => account.key === key
  )?.config;
  const merged: Config = { ...(existing ?? { skypeToken: "" }), ...config };
  if (!config.skypeToken && existing?.skypeToken) {
    merged.skypeToken = existing.skypeToken;
  }
  saveAccount(key, merged, true);

  console.log(`Login successful. Active account: ${key}`);
  if (!merged.skypeToken) {
    console.warn(
      "Warning: no Teams skypetoken was issued for this tenant " +
        "(common for guest/B2B accounts). Teams chat may be unavailable; " +
        "mail/calendar can still work via the refresh token."
    );
  }

  // Auto-discover and register any other tenants this user can access.
  if (!opts.tenant) {
    console.log("\nDiscovering other tenants...");
    try {
      await syncTenants();
    } catch (error) {
      console.warn("Tenant discovery skipped:", (error as Error).message);
    }
  }
  tokenStatus();
}

/** List all stored accounts, marking the current one. */
export function authList(): void {
  const accounts = listAccounts();
  if (accounts.length === 0) {
    console.log("No accounts. Run: ms-cli auth login");
    return;
  }
  const now = Math.floor(Date.now() / 1000);
  for (const { key, current, config } of accounts) {
    const marker = current ? "*" : " ";
    const upn = accountUpn(config);
    let state = "no token";
    try {
      const exp = decodeJwt(config.skypeToken).exp ?? 0;
      state = exp > now ? "valid" : "expired";
    } catch {}
    const parts = [config.tenantName, upn, config.tenantId, state].filter(Boolean);
    console.log(`${marker} ${key}  (${parts.join(" | ")})`);
  }
  console.log(`\nconfig: ${getConfigPath()}`);
}

/** Switch the active account. */
export function authUse(query: string): void {
  const key = useAccount(query);
  console.log(`Switched to account: ${key}`);
  tokenStatus();
}

/** Remove a stored account. */
export function authRemove(key: string): void {
  removeAccount(key);
  console.log(`Removed account: ${key}`);
  const current = getCurrentAccountKey();
  if (current) console.log(`Active account is now: ${current}`);
}
