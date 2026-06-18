import { readFileSync, writeFileSync, mkdirSync, existsSync } from "fs";
import { homedir } from "os";
import { join } from "path";

const CONFIG_DIR = join(homedir(), ".ms-cli");
const CONFIG_FILE = join(CONFIG_DIR, "config.json");

export interface Config {
  skypeToken: string;
  refreshToken?: string;
  refreshTokenIssuedAt?: number; // unix timestamp (seconds)
  outlookToken?: string;
  graphToken?: string;
  formsToken?: string;
  formsTokenExp?: number;
  formsUserId?: string;
  tenantId?: string;
  tenantName?: string; // friendly org name (e.g. "丸紅株式会社(日本国内)")
  userType?: string; // "member" | "guest" — guests have no mailbox in that tenant
  region?: string; // e.g. "jp"
  chatServiceHost?: string;
  clientId?: string; // AAD client that minted refreshToken (default: Teams native client)
}

interface Store {
  version: number;
  current?: string;
  accounts: Record<string, Config>;
}

const STORE_VERSION = 2;
const DEFAULT_CONFIG: Config = { skypeToken: "" };

function emptyStore(): Store {
  return { version: STORE_VERSION, accounts: {} };
}

function readStore(): Store {
  if (!existsSync(CONFIG_FILE)) return emptyStore();
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(CONFIG_FILE, "utf-8"));
  } catch {
    return emptyStore();
  }
  if (!parsed || typeof parsed !== "object") return emptyStore();
  const obj = parsed as Record<string, unknown>;

  // New multi-account format
  if (obj.accounts && typeof obj.accounts === "object") {
    return {
      version: STORE_VERSION,
      current: typeof obj.current === "string" ? obj.current : undefined,
      accounts: obj.accounts as Record<string, Config>,
    };
  }

  // Legacy flat format -> migrate in place and persist
  if (typeof obj.skypeToken === "string") {
    const legacy = obj as unknown as Config;
    const key = legacyAccountKey(legacy);
    const store: Store = { version: STORE_VERSION, current: key, accounts: { [key]: legacy } };
    writeStore(store);
    return store;
  }

  return emptyStore();
}

/** Extract upn/unique_name from a JWT without verifying it (best-effort). */
function jwtUpn(token: string | undefined): string | undefined {
  if (!token) return undefined;
  const parts = token.split(".");
  if (parts.length !== 3) return undefined;
  try {
    const padded = parts[1] + "=".repeat((4 - (parts[1].length % 4)) % 4);
    const payload = JSON.parse(Buffer.from(padded, "base64url").toString());
    return payload.upn ?? payload.unique_name ?? undefined;
  } catch {
    return undefined;
  }
}

/** Friendly key for a migrated legacy account: email if decodable, else tenant. */
function legacyAccountKey(config: Config): string {
  return (
    jwtUpn(config.outlookToken) ??
    jwtUpn(config.graphToken) ??
    config.tenantId ??
    "default"
  );
}

function writeStore(store: Store): void {
  mkdirSync(CONFIG_DIR, { recursive: true });
  writeFileSync(CONFIG_FILE, JSON.stringify(store, null, 2));
}

/**
 * Resolve a user-supplied account reference to an exact account key. Matches,
 * in order: exact key, exact tenantId, tenantId prefix, then case-insensitive
 * substring of the key or tenant name. Throws if nothing or several match.
 */
export function resolveAccountKey(store: Store, query: string): string {
  const keys = Object.keys(store.accounts);
  if (store.accounts[query]) return query;

  const byTenant = keys.filter((k) => store.accounts[k].tenantId === query);
  if (byTenant.length === 1) return byTenant[0];

  const q = query.toLowerCase();
  const fuzzy = keys.filter((k) => {
    const a = store.accounts[k];
    return (
      (a.tenantId ?? "").toLowerCase().startsWith(q) ||
      k.toLowerCase().includes(q) ||
      (a.tenantName ?? "").toLowerCase().includes(q)
    );
  });
  if (fuzzy.length === 1) return fuzzy[0];
  if (fuzzy.length > 1) {
    throw new Error(`"${query}" matches multiple accounts: ${fuzzy.join(", ")}`);
  }
  throw new Error(`Account "${query}" not found. Run: ms-cli auth list`);
}

/**
 * Resolve which account is active. MS_CLI_ACCOUNT overrides the persisted
 * current account for one-off use. An env value that names an unknown account
 * is a hard error (no silent fallback to the wrong account).
 */
function resolveCurrentKey(store: Store): string | undefined {
  const env = process.env.MS_CLI_ACCOUNT;
  if (env) return resolveAccountKey(store, env);
  return store.current;
}

export function loadConfig(): Config {
  const store = readStore();
  const key = resolveCurrentKey(store);
  if (key && store.accounts[key]) return { ...DEFAULT_CONFIG, ...store.accounts[key] };
  return { ...DEFAULT_CONFIG };
}

export function saveConfig(config: Config): void {
  const store = readStore();
  let key = resolveCurrentKey(store);
  if (!key) {
    // First-ever save with no current account: derive a key and make it current.
    key = config.tenantId || "default";
    store.current = key;
  }
  store.accounts[key] = config;
  writeStore(store);
}

export interface AccountEntry {
  key: string;
  current: boolean;
  config: Config;
}

export function listAccounts(): AccountEntry[] {
  const store = readStore();
  return Object.entries(store.accounts).map(([key, config]) => ({
    key,
    current: key === store.current,
    config,
  }));
}

export function getCurrentAccountKey(): string | undefined {
  return resolveCurrentKey(readStore());
}

/** Store (or overwrite) an account under the given key, optionally switching to it. */
export function saveAccount(key: string, config: Config, makeCurrent = true): void {
  const store = readStore();
  store.accounts[key] = config;
  if (makeCurrent) store.current = key;
  writeStore(store);
}

export function useAccount(query: string): string {
  const store = readStore();
  const key = resolveAccountKey(store, query);
  store.current = key;
  writeStore(store);
  return key;
}

export function removeAccount(key: string): void {
  const store = readStore();
  if (!store.accounts[key]) {
    throw new Error(`Account "${key}" not found. Run: ms-cli auth list`);
  }
  delete store.accounts[key];
  if (store.current === key) {
    store.current = Object.keys(store.accounts)[0];
  }
  writeStore(store);
}

export function getConfigPath(): string {
  return CONFIG_FILE;
}
