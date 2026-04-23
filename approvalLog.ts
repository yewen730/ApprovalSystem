import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/** One line, safe for log files (no raw newlines). */
export function approvalLogSanitize(text: string, maxLen = 400): string {
  return String(text ?? "")
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLen);
}

function defaultLogPath(): string {
  const fromEnv = process.env.APPROVAL_EVENT_LOG_PATH?.trim();
  if (fromEnv) return path.isAbsolute(fromEnv) ? fromEnv : path.join(__dirname, fromEnv);
  return path.join(__dirname, "logs", "approval-events.txt");
}

/** Append a single line to the approval / workflow audit log (UTF-8). */
export function approvalEventLog(message: string): void {
  const line = `[${new Date().toISOString()}] ${message}\n`;
  const logPath = defaultLogPath();
  try {
    fs.mkdirSync(path.dirname(logPath), { recursive: true });
    fs.appendFileSync(logPath, line, { encoding: "utf8" });
  } catch (e) {
    console.warn("approvalEventLog write failed:", (e as Error)?.message || e);
  }
}

export function formatActor(req: { user?: { id?: number; username?: string }; entityContext?: string }): string {
  const u = req.user;
  const userPart = u ? `user=${approvalLogSanitize(String(u.username || ""), 120)} id=${u.id ?? "?"}` : "user=?";
  const ent = approvalLogSanitize(String(req.entityContext ?? ""), 32);
  return `${userPart} entity=${ent || "-"}`;
}
