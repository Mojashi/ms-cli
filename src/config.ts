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
  tenantId?: string;
  region?: string; // e.g. "jp"
  chatServiceHost?: string;
}

const DEFAULT_CONFIG: Config = {
  skypeToken: "",
};

export function loadConfig(): Config {
  if (!existsSync(CONFIG_FILE)) return { ...DEFAULT_CONFIG };
  try {
    return { ...DEFAULT_CONFIG, ...JSON.parse(readFileSync(CONFIG_FILE, "utf-8")) };
  } catch {
    return { ...DEFAULT_CONFIG };
  }
}

export function saveConfig(config: Config): void {
  mkdirSync(CONFIG_DIR, { recursive: true });
  writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2));
}

export function getConfigPath(): string {
  return CONFIG_FILE;
}
