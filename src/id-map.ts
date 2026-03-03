import { createHash } from "crypto";
import { readFileSync, writeFileSync, mkdirSync } from "fs";
import { homedir } from "os";
import { join } from "path";

const MAP_FILE = join(homedir(), ".ms-cli", "id-map.json");
const PREFIX_LEN = 8;

interface IdMap {
  [hash: string]: string;
}

/** Generate a short hash from a long ID */
export function shortId(longId: string): string {
  return createHash("sha256").update(longId).digest("hex").slice(0, PREFIX_LEN);
}

/** Register a long ID (saves mapping for later resolution) */
export function registerId(longId: string): string {
  const hash = shortId(longId);
  const map = load();
  map[hash] = longId;
  save(map);
  return hash;
}

/** Batch register and flush at once */
export function registerIds(longIds: string[]): void {
  const map = load();
  for (const id of longIds) {
    map[shortId(id)] = id;
  }
  save(map);
}

/** Resolve a short hash to the full ID. If not a short hash, return as-is. */
export function resolveId(idOrHash: string): string {
  if (idOrHash.length <= PREFIX_LEN && /^[0-9a-f]+$/.test(idOrHash)) {
    const map = load();
    const resolved = map[idOrHash];
    if (!resolved) {
      console.error(`Unknown id: ${idOrHash}`);
      process.exit(1);
    }
    return resolved;
  }
  return idOrHash;
}

function load(): IdMap {
  try {
    return JSON.parse(readFileSync(MAP_FILE, "utf-8"));
  } catch {
    return {};
  }
}

function save(map: IdMap): void {
  mkdirSync(join(homedir(), ".ms-cli"), { recursive: true });
  writeFileSync(MAP_FILE, JSON.stringify(map));
}
