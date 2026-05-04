import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const MAX_ATTACHMENT_BYTES = Number(process.env.MAX_ATTACHMENT_BYTES || 15 * 1024 * 1024);

/** Default UNC root for request attachments and related on-disk files (one folder per request under `{entity}/{department}/{formatted_id}/`). */
export const COMPANY_FILE_STORAGE_ROOT =
  "\\\\10.128.3.10\\data\\E_IVOICING\\Approval System";

/** Root directory on the company server where request files are stored (absolute path recommended). */
export function getAttachmentsRoot(): string {
  const raw = process.env.ATTACHMENTS_STORAGE_PATH?.trim();
  if (raw) return path.isAbsolute(raw) ? raw : path.join(__dirname, raw);
  return COMPANY_FILE_STORAGE_ROOT;
}

function safeSegment(s: string, max = 64): string {
  return String(s || "default")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .slice(0, max);
}

function safeBaseName(originalName: string, max = 80): string {
  const ext = path.extname(originalName || "");
  return (
    path
      .basename(originalName || "file", ext)
      .replace(/[^a-zA-Z0-9._-]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, max) || "file"
  );
}

/** Windows / UNC reserved characters only — keeps preparer spelling, spaces, and Unicode in the basename. */
const FILENAME_FORBIDDEN_CHARS = /[<>:"/\\|?*\u0000-\u001f]/g;

/**
 * Visible `file_name` for DB and downloads: basename from upload, strip path, only remove illegal
 * filename characters (do not replace spaces or unicode with underscores).
 */
export function displayFileNameFromUpload(originalName: string, extLower: string): string {
  const ext = extLower.startsWith(".") ? extLower.toLowerCase() : `.${extLower.toLowerCase()}`;
  const trimmed = String(originalName || "").trim();
  let base = path.basename(trimmed, path.extname(trimmed));
  base = base.replace(FILENAME_FORBIDDEN_CHARS, "_").replace(/[. ]+$/g, "").trim();
  if (!base) base = "file";
  const maxBase = 220;
  return `${base.slice(0, maxBase)}${ext}`;
}

function detectMagic(buffer: Buffer): "pdf" | "png" | "jpg" | "zip" | "ole" | "unknown" {
  if (buffer.length >= 5 && buffer.slice(0, 5).equals(Buffer.from([0x25, 0x50, 0x44, 0x46, 0x2d]))) return "pdf"; // %PDF-
  if (buffer.length >= 8 && buffer.slice(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])))
    return "png";
  if (buffer.length >= 3 && buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff) return "jpg";
  if (
    buffer.length >= 4 &&
    buffer[0] === 0x50 &&
    buffer[1] === 0x4b &&
    (buffer[2] === 0x03 || buffer[2] === 0x05 || buffer[2] === 0x07) &&
    (buffer[3] === 0x04 || buffer[3] === 0x06 || buffer[3] === 0x08)
  )
    return "zip";
  if (
    buffer.length >= 8 &&
    buffer.slice(0, 8).equals(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]))
  )
    return "ole";
  return "unknown";
}

function looksLikeText(buffer: Buffer): boolean {
  // Reject binary-ish content for text formats.
  if (!buffer.length) return false;
  const sample = buffer.subarray(0, Math.min(buffer.length, 4096));
  if (sample.includes(0x00)) return false;
  let nonPrintable = 0;
  for (const b of sample) {
    const isTabOrNewLine = b === 0x09 || b === 0x0a || b === 0x0d;
    const isPrintableAscii = b >= 0x20 && b <= 0x7e;
    if (!isTabOrNewLine && !isPrintableAscii && b < 0x80) nonPrintable++;
  }
  return nonPrintable / sample.length < 0.03;
}

const ALLOWED_ATTACHMENT_TYPES: Record<
  string,
  { mime: string; allowedMagic: Array<ReturnType<typeof detectMagic> | "text">; aliases?: string[] }
> = {
  ".pdf": { mime: "application/pdf", allowedMagic: ["pdf"] },
  ".png": { mime: "image/png", allowedMagic: ["png"] },
  ".jpg": { mime: "image/jpeg", allowedMagic: ["jpg"], aliases: ["image/pjpeg"] },
  ".jpeg": { mime: "image/jpeg", allowedMagic: ["jpg"], aliases: ["image/pjpeg"] },
  ".doc": { mime: "application/msword", allowedMagic: ["ole"] },
  ".docx": {
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    allowedMagic: ["zip"],
  },
  ".xlsx": { mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", allowedMagic: ["zip"] },
  ".csv": { mime: "text/csv", allowedMagic: ["text"], aliases: ["application/csv", "application/vnd.ms-excel"] },
  ".txt": { mime: "text/plain", allowedMagic: ["text"] },
};

export function validateAttachmentUpload(
  originalName: string,
  claimedMimeType: string | null | undefined,
  buffer: Buffer
): { safeFileName: string; displayFileName: string; fileExt: string; mimeType: string } {
  if (!buffer.length) throw new Error("Empty file upload");
  if (!Number.isFinite(MAX_ATTACHMENT_BYTES) || MAX_ATTACHMENT_BYTES <= 0) {
    throw new Error("MAX_ATTACHMENT_BYTES must be a positive integer");
  }
  if (buffer.length > MAX_ATTACHMENT_BYTES) {
    throw new Error(`Attachment exceeds ${MAX_ATTACHMENT_BYTES} bytes limit`);
  }

  const ext = path.extname(String(originalName || "").trim()).toLowerCase();
  const cfg = ALLOWED_ATTACHMENT_TYPES[ext];
  if (!cfg) {
    throw new Error("Unsupported file type. Allowed: PDF, PNG, JPG, DOC, DOCX, XLSX, CSV, TXT");
  }

  const magic = detectMagic(buffer);
  const textOk = looksLikeText(buffer);
  const magicOk = cfg.allowedMagic.some((m) => (m === "text" ? textOk : magic === m));
  if (!magicOk) {
    throw new Error("File content does not match its extension");
  }

  const claimed = String(claimedMimeType || "").trim().toLowerCase();
  if (claimed && claimed !== cfg.mime && !(cfg.aliases || []).includes(claimed)) {
    throw new Error("Claimed MIME type does not match allowed type");
  }

  const safeName = safeBaseName(originalName);
  const displayFileName = displayFileNameFromUpload(originalName, ext);
  return {
    safeFileName: `${safeName}${ext}`,
    displayFileName,
    fileExt: ext,
    mimeType: cfg.mime,
  };
}

/** Decode `data:mime;base64,...` or raw base64 into a buffer. */
export function decodeAttachmentPayload(data: string): Buffer {
  const s = String(data || "").trim();
  if (!s) return Buffer.alloc(0);
  const m = /^data:([^;]*);base64,(.*)$/s.exec(s);
  if (m) return Buffer.from(m[2], "base64");
  return Buffer.from(s, "base64");
}

/**
 * Resolve a path stored in `request_attachments.file_path`: supports full UNC/absolute paths
 * (new) and legacy paths relative to `getAttachmentsRoot()`.
 */
export function resolveStoredPath(storedPath: string): string {
  const raw = String(storedPath || "").trim();
  if (!raw || raw.includes("..")) throw new Error("Invalid stored path");
  const rootRaw = getAttachmentsRoot();
  const rootResolved = path.resolve(rootRaw);

  const normalizedInput = raw.replace(/\//g, path.sep);
  let candidate: string;
  if (path.isAbsolute(normalizedInput)) {
    candidate = path.resolve(normalizedInput);
  } else {
    const rel = normalizedInput.replace(/^[\\/]+/, "");
    candidate = path.resolve(path.join(rootResolved, ...rel.split(/[/\\]+/)));
  }

  const rootLc = rootResolved.toLowerCase();
  const candLc = candidate.toLowerCase();
  const rootPrefix = rootLc.endsWith(path.sep) ? rootLc : rootLc + path.sep;
  if (candLc !== rootLc && !candLc.startsWith(rootPrefix)) {
    throw new Error("Path escapes attachments root");
  }
  return candidate;
}

export function saveRequestAttachmentFile(
  entity: string,
  department: string | null | undefined,
  formattedId: string | null | undefined,
  requestId: number,
  originalName: string,
  buffer: Buffer,
  claimedMimeType?: string | null
): { storedPath: string; mimeType: string; storedFileName: string } {
  const validated = validateAttachmentUpload(originalName, claimedMimeType, buffer);
  const folderId =
    String(formattedId || "").trim() || `request_${Number.isFinite(requestId) ? String(requestId) : "unknown"}`;
  const dir = path.join(
    getAttachmentsRoot(),
    safeSegment(entity),
    safeSegment(String(department || "general")),
    safeSegment(folderId, 128)
  );
  fs.mkdirSync(dir, { recursive: true });
  const unique = `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  const filename = `${unique}_${validated.safeFileName}`;
  const full = path.join(dir, filename);
  fs.writeFileSync(full, buffer);
  const storedPath = path.normalize(path.resolve(full));
  return { storedPath, mimeType: validated.mimeType, storedFileName: validated.displayFileName };
}

/**
 * Same folder layout as {@link saveRequestAttachmentFile} (`entity` / `department` / `formatted_id` /),
 * but uses a fixed PDF basename (e.g. `PR_PRGC26-00001.pdf`) and overwrites the previous file.
 * Matches on-disk layout under `COMPANY_FILE_STORAGE_ROOT` (e.g. `…/GCCM/Maintenance/PRGC26-00001/`).
 */
export function saveGeneratedProcurementFormPdf(
  entity: string,
  department: string | null | undefined,
  formattedId: string | null | undefined,
  requestId: number,
  originalName: string,
  buffer: Buffer,
  claimedMimeType?: string | null
): { storedPath: string; mimeType: string; storedFileName: string } {
  const validated = validateAttachmentUpload(originalName, claimedMimeType, buffer);
  const folderId =
    String(formattedId || "").trim() || `request_${Number.isFinite(requestId) ? String(requestId) : "unknown"}`;
  const dir = path.join(
    getAttachmentsRoot(),
    safeSegment(entity),
    safeSegment(String(department || "general")),
    safeSegment(folderId, 128)
  );
  fs.mkdirSync(dir, { recursive: true });
  const full = path.join(dir, validated.safeFileName);
  try {
    if (fs.existsSync(full)) fs.unlinkSync(full);
  } catch {
    // ignore
  }
  fs.writeFileSync(full, buffer);
  const storedPath = path.normalize(path.resolve(full));
  return { storedPath, mimeType: validated.mimeType, storedFileName: validated.displayFileName };
}

export function tryUnlinkStoredFile(storedPath: string | null | undefined): void {
  if (!storedPath) return;
  try {
    const full = resolveStoredPath(storedPath);
    if (fs.existsSync(full)) fs.unlinkSync(full);
  } catch {
    // ignore
  }
}

export function copyStoredFileToRequest(
  entity: string,
  department: string | null | undefined,
  formattedId: string | null | undefined,
  newRequestId: number,
  sourceRelativePath: string | null | undefined,
  originalFileName: string
): { storedPath: string } | null {
  if (!sourceRelativePath) return null;
  try {
    const src = resolveStoredPath(sourceRelativePath);
    if (!fs.existsSync(src)) return null;
    const buf = fs.readFileSync(src);
    return saveRequestAttachmentFile(entity, department, formattedId, newRequestId, originalFileName, buf);
  } catch {
    return null;
  }
}
