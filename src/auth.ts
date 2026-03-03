import { createInterface } from "readline/promises";
import { loadConfig, saveConfig, getConfigPath } from "./config.js";

interface JwtPayload {
  iat?: number;
  exp?: number;
  skypeid?: string;
  scp?: number;
  rgn?: string;
  tid?: string;
  aud?: string;
  [key: string]: unknown;
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

function getClientId(config: { clientId?: string }): string {
  if (!config.clientId) {
    console.error("No clientId configured. Set it in ~/.ms-cli/config.json");
    process.exit(1);
  }
  return config.clientId;
}

/** Try to refresh skypetoken using saved refresh token. Returns true on success. */
export async function tryRefresh(quiet = false): Promise<boolean> {
  const config = loadConfig();
  if (!config.refreshToken) {
    if (!quiet) console.error("No refresh token saved. Run: ms-cli auth login");
    return false;
  }
  const clientId = getClientId(config);

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

/** Device code flow: get refresh token interactively via browser auth */
export async function deviceCodeLogin(): Promise<void> {
  const config = loadConfig();
  const clientId = getClientId(config);
  const tenantId = config.tenantId ?? "common";

  // Step 1: Request device code
  const codeRes = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        scope: "https://api.spaces.skype.com/.default openid profile offline_access",
      }),
    }
  );

  if (!codeRes.ok) {
    const text = await codeRes.text();
    console.error(`Device code request failed (${codeRes.status}): ${text}`);
    process.exit(1);
  }

  const codeData = (await codeRes.json()) as {
    device_code: string;
    user_code: string;
    verification_uri: string;
    expires_in: number;
    interval: number;
    message: string;
  };

  console.log(codeData.message);

  // Open browser automatically
  const { execSync } = await import("child_process");
  try {
    execSync(`open "${codeData.verification_uri}"`);
  } catch {}

  // Step 2: Poll for token
  const interval = (codeData.interval ?? 5) * 1000;
  const deadline = Date.now() + codeData.expires_in * 1000;

  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, interval));

    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: clientId,
          grant_type: "urn:ietf:params:oauth:grant-type:device_code",
          device_code: codeData.device_code,
        }),
      }
    );

    const tokenData = (await tokenRes.json()) as {
      access_token?: string;
      refresh_token?: string;
      error?: string;
    };

    if (tokenData.error === "authorization_pending") continue;
    if (tokenData.error === "slow_down") {
      await new Promise((r) => setTimeout(r, 5000));
      continue;
    }
    if (tokenData.error) {
      console.error(`Auth failed: ${tokenData.error}`);
      process.exit(1);
    }

    // Success
    if (tokenData.refresh_token) {
      config.refreshToken = tokenData.refresh_token;
      config.refreshTokenIssuedAt = Math.floor(Date.now() / 1000);
    }

    // Exchange AAD token for skypetoken
    if (tokenData.access_token) {
      console.log("Got AAD token. Exchanging for skypetoken...");
      for (const endpoint of [
        "https://teams.microsoft.com/api/authsvc/v1.0/authz",
        "https://authsvc.teams.microsoft.com/v1.0/authz",
      ]) {
        try {
          const skypeRes = await fetch(endpoint, {
            method: "POST",
            headers: {
              Authorization: `Bearer ${tokenData.access_token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({}),
          });
          if (skypeRes.ok) {
            const skypeData = (await skypeRes.json()) as { tokens?: { skypeToken?: string } };
            if (skypeData.tokens?.skypeToken) {
              config.skypeToken = skypeData.tokens.skypeToken;
              break;
            }
          }
        } catch {}
      }
    }

    saveConfig(config);
    console.log("Login successful.");
    tokenStatus();
    return;
  }

  console.error("Timed out waiting for authentication.");
  process.exit(1);
}

/** Interactive setup: configure clientId and run device code login */
export async function setup(): Promise<void> {
  const config = loadConfig();

  const rl = createInterface({ input: process.stdin, output: process.stdout });
  try {
    const clientId = await rl.question("Client ID: ");
    if (!clientId.trim()) {
      console.error("Client ID is required.");
      process.exit(1);
    }
    config.clientId = clientId.trim();
    saveConfig(config);
    console.log("Client ID saved.\n");
  } finally {
    rl.close();
  }

  await deviceCodeLogin();
}
