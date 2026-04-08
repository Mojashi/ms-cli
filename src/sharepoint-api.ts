import { writeFileSync, readFileSync, mkdirSync, statSync } from "fs";
import { join, basename } from "path";
import { loadConfig, saveConfig, type Config } from "./config.js";

const TEAMS_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

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

export async function ensureGraphToken(): Promise<string> {
  const config = loadConfig();

  if (config.graphToken) {
    try {
      const payload = JSON.parse(
        Buffer.from(config.graphToken.split(".")[1] + "==", "base64url").toString()
      );
      if (payload.exp && payload.exp > Date.now() / 1000 + 60) {
        return config.graphToken;
      }
    } catch {}
  }

  if (!config.refreshToken) {
    console.error("No refresh token. Run: ms-cli auth login");
    process.exit(1);
  }

  console.error("Refreshing Graph token...");
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
        client_id: TEAMS_CLIENT_ID,
        grant_type: "refresh_token",
        refresh_token: config.refreshToken,
        scope: "https://graph.microsoft.com/.default openid profile offline_access",
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

  config.graphToken = data.access_token;
  if (data.refresh_token) {
    config.refreshToken = data.refresh_token;
    config.refreshTokenIssuedAt = Math.floor(Date.now() / 1000);
  }
  saveConfig(config);
  console.error("Graph token refreshed.");
  return data.access_token;
}

const MAX_RETRIES = 3;

async function graphFetch(url: string, init: RequestInit): Promise<Response> {
  const fullUrl = url.startsWith("http") ? url : `${GRAPH_BASE}${url}`;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const res = await fetch(fullUrl, init);
    if (res.status === 429 || res.status === 503) {
      if (attempt >= MAX_RETRIES) {
        const text = await res.text();
        throw new Error(`Graph API throttled after ${MAX_RETRIES} retries: ${text.slice(0, 300)}`);
      }
      const retryAfter = parseInt(res.headers.get("Retry-After") ?? "", 10);
      const waitSec = Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter : 2 ** attempt + 1;
      console.error(`${c.yellow}Rate limited. Retrying in ${waitSec}s... (${attempt + 1}/${MAX_RETRIES})${c.reset}`);
      await new Promise((r) => setTimeout(r, waitSec * 1000));
      continue;
    }
    return res;
  }
  throw new Error("Unreachable");
}

async function graphGet(path: string): Promise<unknown> {
  const token = await ensureGraphToken();
  const res = await graphFetch(path, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res.json();
}

async function graphGetRaw(path: string): Promise<Response> {
  const token = await ensureGraphToken();
  const res = await graphFetch(path, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API error ${res.status}: ${text.slice(0, 300)}`);
  }
  return res;
}

// --- Types ---

interface ODataResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

interface Site {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  description?: string;
  lastModifiedDateTime?: string;
}

interface Drive {
  id: string;
  name: string;
  driveType: string;
  webUrl: string;
  quota?: {
    total: number;
    used: number;
    remaining: number;
  };
}

interface DriveItem {
  id: string;
  name: string;
  webUrl: string;
  size?: number;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { displayName?: string } };
  folder?: { childCount: number };
  file?: { mimeType: string };
  "@microsoft.graph.downloadUrl"?: string;
}

// --- Sites ---

export async function spSites(options: {
  query?: string;
  pageSize?: number;
}): Promise<void> {
  const pageSize = options.pageSize ?? 20;
  let data: ODataResponse<Site>;

  if (options.query) {
    data = (await graphGet(
      `/sites?search=${encodeURIComponent(options.query)}&$top=${pageSize}`
    )) as ODataResponse<Site>;
  } else {
    // List sites the user follows or has access to
    data = (await graphGet(
      `/sites?search=*&$top=${pageSize}`
    )) as ODataResponse<Site>;
  }

  if (data.value.length === 0) {
    console.log("No sites found.");
    return;
  }

  for (const site of data.value) {
    const desc = site.description ? ` ${c.dim}${site.description.slice(0, 60)}${c.reset}` : "";
    console.log(`${c.bold}${c.cyan}${site.displayName || site.name}${c.reset}${desc}`);
    console.log(`  ${c.dim}id: ${site.id}${c.reset}`);
    console.log(`  ${c.dim}url: ${site.webUrl}${c.reset}`);
    console.log();
  }

  console.log(`${c.bold}${data.value.length}${c.reset} sites`);
}

// --- Drives (Document Libraries) ---

export async function spDrives(siteId: string): Promise<void> {
  const data = (await graphGet(
    `/sites/${encodeURIComponent(siteId)}/drives`
  )) as ODataResponse<Drive>;

  if (data.value.length === 0) {
    console.log("No drives found.");
    return;
  }

  for (const drive of data.value) {
    const quota = drive.quota
      ? ` ${c.dim}(${formatSize(drive.quota.used)} / ${formatSize(drive.quota.total)})${c.reset}`
      : "";
    console.log(`${c.bold}${c.blue}${drive.name}${c.reset} ${c.dim}[${drive.driveType}]${c.reset}${quota}`);
    console.log(`  ${c.dim}id: ${drive.id}${c.reset}`);
    console.log(`  ${c.dim}url: ${drive.webUrl}${c.reset}`);
    console.log();
  }

  console.log(`${c.bold}${data.value.length}${c.reset} drives`);
}

// --- List Files ---

export async function spFiles(
  driveId: string,
  options: { path?: string; pageSize?: number; json?: boolean }
): Promise<void> {
  const pageSize = options.pageSize ?? 30;
  const itemPath = options.path ? `:/${encodeURIComponent(options.path)}:` : "/root";

  const data = (await graphGet(
    `/drives/${encodeURIComponent(driveId)}${itemPath}/children?$top=${pageSize}&$orderby=name`
  )) as ODataResponse<DriveItem>;

  if (options.json) {
    console.log(JSON.stringify(data.value, null, 2));
    return;
  }

  if (data.value.length === 0) {
    console.log("Empty folder.");
    return;
  }

  for (const item of data.value) {
    const isFolder = !!item.folder;
    const icon = isFolder ? `${c.blue}📁` : `${c.white}📄`;
    const sizeStr = item.size != null && !isFolder ? ` ${c.dim}${formatSize(item.size)}${c.reset}` : "";
    const childCount = isFolder && item.folder ? ` ${c.dim}(${item.folder.childCount} items)${c.reset}` : "";
    const modified = item.lastModifiedDateTime ? formatTime(item.lastModifiedDateTime) : "";
    const modifiedBy = item.lastModifiedBy?.user?.displayName ?? "";

    console.log(`${icon} ${c.bold}${item.name}${c.reset}${sizeStr}${childCount}`);
    console.log(`  ${c.dim}id: ${item.id}${c.reset}`);
    if (modified || modifiedBy) {
      console.log(`  ${c.dim}${modified}${modifiedBy ? ` by ${modifiedBy}` : ""}${c.reset}`);
    }
    console.log();
  }

  console.log(`${c.bold}${data.value.length}${c.reset} items`);
}

// --- Download ---

export async function spDownload(
  driveId: string,
  itemId: string,
  options: { outDir?: string }
): Promise<void> {
  // Get item metadata first
  const item = (await graphGet(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`
  )) as DriveItem;

  if (item.folder) {
    console.error("Cannot download a folder. Use sp files to browse its contents.");
    process.exit(1);
  }

  const downloadUrl = item["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) {
    console.error("No download URL available for this item.");
    process.exit(1);
  }

  console.log(`Downloading ${c.bold}${item.name}${c.reset} (${formatSize(item.size ?? 0)})...`);

  const res = await fetch(downloadUrl);
  if (!res.ok) {
    console.error(`Download failed: ${res.status}`);
    process.exit(1);
  }

  const outDir = options.outDir ?? ".";
  mkdirSync(outDir, { recursive: true });
  const outPath = join(outDir, item.name);
  writeFileSync(outPath, Buffer.from(await res.arrayBuffer()));

  console.log(`${c.green}Saved:${c.reset} ${outPath}`);
}

// --- Search ---

export async function spSearch(
  query: string,
  options: { pageSize?: number }
): Promise<void> {
  const pageSize = options.pageSize ?? 15;

  // Use Graph search API
  const token = await ensureGraphToken();
  const res = await fetch(`${GRAPH_BASE}/search/query`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ["driveItem"],
          query: { queryString: query },
          from: 0,
          size: pageSize,
        },
      ],
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Search API error ${res.status}: ${text.slice(0, 300)}`);
  }

  const data = (await res.json()) as {
    value: Array<{
      hitsContainers: Array<{
        total: number;
        hits: Array<{
          summary?: string;
          resource: {
            id: string;
            name: string;
            webUrl: string;
            size?: number;
            lastModifiedDateTime?: string;
            lastModifiedBy?: { user?: { displayName?: string } };
            parentReference?: { driveId?: string; siteId?: string };
          };
        }>;
      }>;
    }>;
  };

  const container = data.value?.[0]?.hitsContainers?.[0];
  if (!container || !container.hits || container.hits.length === 0) {
    console.log("No results found.");
    return;
  }

  for (const hit of container.hits) {
    const r = hit.resource;
    const sizeStr = r.size != null ? ` ${c.dim}${formatSize(r.size)}${c.reset}` : "";
    const modified = r.lastModifiedDateTime ? formatTime(r.lastModifiedDateTime) : "";
    const modifiedBy = r.lastModifiedBy?.user?.displayName ?? "";
    const summary = hit.summary
      ? `  ${c.dim}${hit.summary.replace(/<\/?[^>]+>/g, "").slice(0, 80)}${c.reset}`
      : "";

    console.log(`${c.yellow}📄${c.reset} ${c.bold}${r.name}${c.reset}${sizeStr}`);
    console.log(`  ${c.dim}id: ${r.id}${c.reset}`);
    if (r.parentReference?.driveId) {
      console.log(`  ${c.dim}drive: ${r.parentReference.driveId}${c.reset}`);
    }
    console.log(`  ${c.dim}url: ${r.webUrl}${c.reset}`);
    if (modified || modifiedBy) {
      console.log(`  ${c.dim}${modified}${modifiedBy ? ` by ${modifiedBy}` : ""}${c.reset}`);
    }
    if (summary) console.log(summary);
    console.log();
  }

  console.log(`${c.bold}${container.hits.length}${c.reset} / ${container.total} results`);
}

// --- Recent files ---

export async function spRecent(options: { pageSize?: number }): Promise<void> {
  const pageSize = options.pageSize ?? 20;

  const data = (await graphGet(
    `/me/drive/recent?$top=${pageSize}`
  )) as ODataResponse<DriveItem>;

  if (data.value.length === 0) {
    console.log("No recent files.");
    return;
  }

  for (const item of data.value) {
    const isFolder = !!item.folder;
    const icon = isFolder ? `${c.blue}📁` : `${c.yellow}📄`;
    const sizeStr = item.size != null && !isFolder ? ` ${c.dim}${formatSize(item.size)}${c.reset}` : "";
    const modified = item.lastModifiedDateTime ? formatTime(item.lastModifiedDateTime) : "";
    const modifiedBy = item.lastModifiedBy?.user?.displayName ?? "";

    console.log(`${icon} ${c.bold}${item.name}${c.reset}${sizeStr}`);
    console.log(`  ${c.dim}${modified}${modifiedBy ? ` by ${modifiedBy}` : ""}${c.reset}`);
    console.log(`  ${c.dim}url: ${item.webUrl}${c.reset}`);
    console.log();
  }

  console.log(`${c.bold}${data.value.length}${c.reset} recent files`);
}

// --- Open in browser ---

export async function spOpen(
  driveId: string,
  itemId: string
): Promise<void> {
  const item = (await graphGet(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=name,webUrl`
  )) as DriveItem;

  console.log(`${c.bold}Opening:${c.reset} ${item.name}`);
  const { execSync } = await import("child_process");
  execSync(`open ${JSON.stringify(item.webUrl)}`);
}

// --- Upload ---

const SMALL_FILE_LIMIT = 4 * 1024 * 1024; // 4 MB

export async function spUpload(
  driveId: string,
  localPath: string,
  options: { remotePath?: string }
): Promise<void> {
  const fileName = basename(localPath);
  const remoteName = options.remotePath ?? fileName;
  const fileSize = statSync(localPath).size;
  const fileData = readFileSync(localPath);

  const token = await ensureGraphToken();

  if (fileSize <= SMALL_FILE_LIMIT) {
    // Simple upload for small files
    console.log(`Uploading ${c.bold}${fileName}${c.reset} (${formatSize(fileSize)})...`);
    const res = await graphFetch(
      `/drives/${encodeURIComponent(driveId)}/root:/${encodeURIComponent(remoteName)}:/content`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/octet-stream",
        },
        body: fileData,
      }
    );
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Upload failed (${res.status}): ${text.slice(0, 300)}`);
    }
    const item = (await res.json()) as DriveItem;
    console.log(`${c.green}Uploaded:${c.reset} ${item.name}`);
    console.log(`  ${c.dim}id: ${item.id}${c.reset}`);
    console.log(`  ${c.dim}url: ${item.webUrl}${c.reset}`);
  } else {
    // Upload session for large files
    console.log(`Uploading ${c.bold}${fileName}${c.reset} (${formatSize(fileSize)}) via upload session...`);
    const sessionRes = await graphFetch(
      `/drives/${encodeURIComponent(driveId)}/root:/${encodeURIComponent(remoteName)}:/createUploadSession`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({}),
      }
    );
    if (!sessionRes.ok) {
      const text = await sessionRes.text();
      throw new Error(`Create upload session failed (${sessionRes.status}): ${text.slice(0, 300)}`);
    }
    const session = (await sessionRes.json()) as { uploadUrl: string };

    const chunkSize = 10 * 1024 * 1024; // 10 MB chunks
    let offset = 0;
    let item: DriveItem | undefined;

    while (offset < fileSize) {
      const end = Math.min(offset + chunkSize, fileSize);
      const chunk = fileData.subarray(offset, end);
      const contentRange = `bytes ${offset}-${end - 1}/${fileSize}`;

      const chunkRes = await fetch(session.uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": chunk.length.toString(),
          "Content-Range": contentRange,
        },
        body: chunk,
      });

      if (!chunkRes.ok) {
        const text = await chunkRes.text();
        throw new Error(`Chunk upload failed at ${contentRange} (${chunkRes.status}): ${text.slice(0, 300)}`);
      }

      offset = end;
      const pct = Math.round((offset / fileSize) * 100);
      console.log(`  ${c.dim}${pct}%${c.reset}`);

      if (chunkRes.status === 200 || chunkRes.status === 201) {
        item = (await chunkRes.json()) as DriveItem;
      }
    }

    if (item) {
      console.log(`${c.green}Uploaded:${c.reset} ${item.name}`);
      console.log(`  ${c.dim}id: ${item.id}${c.reset}`);
      console.log(`  ${c.dim}url: ${item.webUrl}${c.reset}`);
    } else {
      console.log(`${c.green}Upload complete.${c.reset}`);
    }
  }
}

// --- Delete ---

export async function spDelete(
  driveId: string,
  itemId: string
): Promise<void> {
  const item = (await graphGet(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=name,size`
  )) as DriveItem;

  console.log(`Deleting ${c.bold}${item.name}${c.reset}...`);

  const token = await ensureGraphToken();
  const res = await graphFetch(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`,
    {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Delete failed (${res.status}): ${text.slice(0, 300)}`);
  }

  console.log(`${c.green}Deleted:${c.reset} ${item.name}`);
}

// --- Convert (download as different format) ---

const SUPPORTED_FORMATS = ["pdf", "html", "jpg", "png", "glb"] as const;
type ConvertFormat = (typeof SUPPORTED_FORMATS)[number];

export async function spConvert(
  driveId: string,
  itemId: string,
  options: { format?: string; outDir?: string }
): Promise<void> {
  const format = (options.format ?? "pdf") as ConvertFormat;
  if (!SUPPORTED_FORMATS.includes(format)) {
    console.error(`Unsupported format: ${format}. Supported: ${SUPPORTED_FORMATS.join(", ")}`);
    process.exit(1);
  }

  // Get item metadata for the filename
  const item = (await graphGet(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=name,size`
  )) as DriveItem;

  const outName = item.name.replace(/\.[^.]+$/, `.${format}`);
  console.log(`Converting ${c.bold}${item.name}${c.reset} → ${c.bold}${outName}${c.reset}...`);

  const res = await graphGetRaw(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content?format=${format}`
  );

  const outDir = options.outDir ?? ".";
  mkdirSync(outDir, { recursive: true });
  const outPath = join(outDir, outName);
  writeFileSync(outPath, Buffer.from(await res.arrayBuffer()));

  console.log(`${c.green}Saved:${c.reset} ${outPath}`);
}

// --- Upload + Convert (one-shot) ---

export async function spUploadConvert(
  driveId: string,
  localPath: string,
  options: { format?: string; outDir?: string }
): Promise<void> {
  const format = (options.format ?? "pdf") as ConvertFormat;
  if (!SUPPORTED_FORMATS.includes(format)) {
    console.error(`Unsupported format: ${format}. Supported: ${SUPPORTED_FORMATS.join(", ")}`);
    process.exit(1);
  }

  // 1. Upload (use a temp name to avoid locking conflicts with existing files)
  const fileName = basename(localPath);
  const ext = fileName.includes(".") ? fileName.slice(fileName.lastIndexOf(".")) : "";
  const base = fileName.includes(".") ? fileName.slice(0, fileName.lastIndexOf(".")) : fileName;
  const remoteName = `${base}_convert_${Date.now()}${ext}`;
  const fileSize = statSync(localPath).size;
  const fileData = readFileSync(localPath);
  const token = await ensureGraphToken();

  console.log(`Uploading ${c.bold}${fileName}${c.reset} (${formatSize(fileSize)})...`);

  let item: DriveItem;
  if (fileSize <= SMALL_FILE_LIMIT) {
    const res = await graphFetch(
      `/drives/${encodeURIComponent(driveId)}/root:/${encodeURIComponent(remoteName)}:/content`,
      {
        method: "PUT",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" },
        body: fileData,
      }
    );
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Upload failed (${res.status}): ${text.slice(0, 300)}`);
    }
    item = (await res.json()) as DriveItem;
  } else {
    const sessionRes = await graphFetch(
      `/drives/${encodeURIComponent(driveId)}/root:/${encodeURIComponent(remoteName)}:/createUploadSession`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({}),
      }
    );
    if (!sessionRes.ok) {
      const text = await sessionRes.text();
      throw new Error(`Create upload session failed (${sessionRes.status}): ${text.slice(0, 300)}`);
    }
    const session = (await sessionRes.json()) as { uploadUrl: string };
    const chunkSize = 10 * 1024 * 1024;
    let offset = 0;
    let uploaded: DriveItem | undefined;
    while (offset < fileSize) {
      const end = Math.min(offset + chunkSize, fileSize);
      const chunk = fileData.subarray(offset, end);
      // Upload session URL is pre-authenticated, no Graph rate limit — use fetch directly
      const chunkRes = await fetch(session.uploadUrl, {
        method: "PUT",
        headers: { "Content-Length": chunk.length.toString(), "Content-Range": `bytes ${offset}-${end - 1}/${fileSize}` },
        body: chunk,
      });
      if (!chunkRes.ok) {
        const text = await chunkRes.text();
        throw new Error(`Chunk upload failed (${chunkRes.status}): ${text.slice(0, 300)}`);
      }
      offset = end;
      console.log(`  ${c.dim}${Math.round((offset / fileSize) * 100)}%${c.reset}`);
      if (chunkRes.status === 200 || chunkRes.status === 201) {
        uploaded = (await chunkRes.json()) as DriveItem;
      }
    }
    if (!uploaded) throw new Error("Upload completed but no item returned");
    item = uploaded;
  }

  console.log(`${c.green}Uploaded.${c.reset} Converting to ${format}...`);

  // 2. Convert + download
  const outName = fileName.replace(/\.[^.]+$/, `.${format}`);
  const res = await graphGetRaw(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(item.id)}/content?format=${format}`
  );
  const outDir = options.outDir ?? ".";
  mkdirSync(outDir, { recursive: true });
  const outPath = join(outDir, outName);
  writeFileSync(outPath, Buffer.from(await res.arrayBuffer()));
  console.log(`${c.green}Saved:${c.reset} ${outPath}`);

  // 3. Cleanup remote
  await spDelete(driveId, item.id);
}

// --- Helpers ---

function formatSize(bytes: number): string {
  if (bytes >= 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024 / 1024).toFixed(1)} GB`;
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  if (bytes >= 1024) return `${Math.ceil(bytes / 1024)} KB`;
  return `${bytes} B`;
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
