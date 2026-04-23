import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/** Root directory on the company server where request files are stored (absolute path recommended). */
export function getAttachmentsRoot(): string {
  const raw = process.env.ATTACHMENTS_STORAGE_PATH?.trim();
  if (raw) return path.isAbsolute(raw) ? raw : path.join(__dirname, raw);
  return path.join(__dirname, "data", "request-attachments");
}

function safeSegment(s: string, max = 64): string {
  return String(s || "default")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .slice(0, max);
}

/** Decode `data:mime;base64,...` or raw base64 into a buffer. */
export function decodeAttachmentPayload(data: string): Buffer {
  const s = String(data || "").trim();
  if (!s) return Buffer.alloc(0);
  const m = /^data:([^;]*);base64,(.*)$/s.exec(s);
  if (m) return Buffer.from(m[2], "base64");
  return Buffer.from(s, "base64");
}

export function resolveStoredPath(relativePath: string): string {
  const rel = String(relativePath || "").replace(/\\/g, "/").replace(/^\/+/, "");
  if (!rel || rel.includes("..")) throw new Error("Invalid stored path");
  const full = path.join(getAttachmentsRoot(), ...rel.split("/"));
  const root = path.resolve(getAttachmentsRoot());
  const resolved = path.resolve(full);
  if (!resolved.startsWith(root + path.sep) && resolved !== root) {
    throw new Error("Path escapes attachments root");
  }
  return resolved;
}

export function saveRequestAttachmentFile(
  entity: string,
  requestId: number,
  originalName: string,
  buffer: Buffer
): { relativePath: string } {
  if (!buffer.length) throw new Error("Empty file upload");
  const dir = path.join(getAttachmentsRoot(), safeSegment(entity), String(requestId));
  fs.mkdirSync(dir, { recursive: true });
  const ext = path.extname(originalName) || "";
  const base = path.basename(originalName, ext).replace(/[^a-zA-Z0-9._-]+/g, "_").slice(0, 80) || "file";
  const unique = `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  const filename = `${unique}_${base}${ext}`;
  const full = path.join(dir, filename);
  fs.writeFileSync(full, buffer);
  const rel = path.relative(getAttachmentsRoot(), full).split(path.sep).join("/");
  return { relativePath: rel };
}

export function tryUnlinkStoredFile(relativePath: string | null | undefined): void {
  if (!relativePath) return;
  try {
    const full = resolveStoredPath(relativePath);
    if (fs.existsSync(full)) fs.unlinkSync(full);
  } catch {
    // ignore
  }
}

export function copyStoredFileToRequest(
  entity: string,
  newRequestId: number,
  sourceRelativePath: string | null | undefined,
  originalFileName: string
): { relativePath: string } | null {
  if (!sourceRelativePath) return null;
  try {
    const src = resolveStoredPath(sourceRelativePath);
    if (!fs.existsSync(src)) return null;
    const buf = fs.readFileSync(src);
    return saveRequestAttachmentFile(entity, newRequestId, originalFileName, buf);
  } catch {
    return null;
  }
}
