/**
 * One-time: copy Flowmaster data from database.db (SQLite) into SQL Server.
 *
 * Prerequisites:
 *   - SQLSERVER_CONNECTION_STRING in .env (same DB the app uses)
 *   - [dbo].[usersetting] rows exist for every username referenced in SQLite
 *     (script maps SQLite user id -> SQL id by matching username)
 *
 * Run:  npx tsx scripts/migrate-sqlite-to-sqlserver.ts
 *
 * Optional env:
 *   SQLITE_PATH=path/to/database.db   (default: ./database.db)
 */
import "dotenv/config";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import Database from "better-sqlite3";
import sql from "mssql";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, "..");
const sqlitePath = process.env.SQLITE_PATH || path.join(root, "database.db");
const conn = process.env.SQLSERVER_CONNECTION_STRING || "";

const DEFAULT_USERS_TABLE = "[dbo].[usersetting]";
const rawUsers = (process.env.SQLSERVER_USERS_TABLE || DEFAULT_USERS_TABLE).trim();
const USERS_TABLE =
  /^[\w\[\].]+$/.test(rawUsers) && rawUsers.length < 200 ? rawUsers : DEFAULT_USERS_TABLE;

async function ensureColumns(pool: sql.ConnectionPool) {
  await pool.request().query(`
    IF COL_LENGTH('request_attachments', 'file_path') IS NULL
      ALTER TABLE request_attachments ADD file_path NVARCHAR(1000) NULL;
    IF COL_LENGTH('request_attachments', 'file_data') IS NOT NULL
      ALTER TABLE request_attachments ALTER COLUMN file_data NVARCHAR(MAX) NULL;
    IF COL_LENGTH('attachments', 'file_data') IS NOT NULL
      ALTER TABLE attachments ALTER COLUMN file_data NVARCHAR(MAX) NULL;
    IF COL_LENGTH('workflow_requests', 'requester_signature') IS NOT NULL
      ALTER TABLE workflow_requests ALTER COLUMN requester_signature NVARCHAR(MAX) NULL;
    IF COL_LENGTH('request_approvals', 'approver_signature') IS NOT NULL
      ALTER TABLE request_approvals ALTER COLUMN approver_signature NVARCHAR(MAX) NULL;
    IF COL_LENGTH('request_approvals', 'approver_username') IS NULL
      ALTER TABLE request_approvals ADD approver_username NVARCHAR(255) NULL;
    IF COL_LENGTH('request_approvals', 'request_title_snapshot') IS NULL
      ALTER TABLE request_approvals ADD request_title_snapshot NVARCHAR(1000) NULL;
    IF COL_LENGTH('request_approvals', 'request_formatted_id_snapshot') IS NULL
      ALTER TABLE request_approvals ADD request_formatted_id_snapshot NVARCHAR(255) NULL;
    IF COL_LENGTH('request_approvals', 'approver_role_snapshot') IS NULL
      ALTER TABLE request_approvals ADD approver_role_snapshot NVARCHAR(255) NULL;
    IF COL_LENGTH('workflow_requests', 'requester_username_snapshot') IS NULL
      ALTER TABLE workflow_requests ADD requester_username_snapshot NVARCHAR(255) NULL;
    IF COL_LENGTH('workflow_requests', 'template_name_snapshot') IS NULL
      ALTER TABLE workflow_requests ADD template_name_snapshot NVARCHAR(255) NULL;
    IF COL_LENGTH('workflow_requests', 'requester_name') IS NULL
      ALTER TABLE workflow_requests ADD requester_name NVARCHAR(255) NULL;
    IF COL_LENGTH('workflow_requests', 'checked_at') IS NULL
      ALTER TABLE workflow_requests ADD checked_at DATETIME2(3) NULL;
    IF COL_LENGTH('workflow_requests', 'checker_name') IS NULL
      ALTER TABLE workflow_requests ADD checker_name NVARCHAR(255) NULL;
    IF COL_LENGTH('workflow_requests', 'approved_at') IS NULL
      ALTER TABLE workflow_requests ADD approved_at DATETIME2(3) NULL;
    IF COL_LENGTH('workflow_requests', 'approver_name') IS NULL
      ALTER TABLE workflow_requests ADD approver_name NVARCHAR(255) NULL;
    IF COL_LENGTH('workflow_requests', 'requester_signed_at') IS NULL
      ALTER TABLE workflow_requests ADD requester_signed_at DATETIME2(3) NULL;
    IF COL_LENGTH('workflow_requests', 'discount_rate') IS NULL
      ALTER TABLE workflow_requests ADD discount_rate DECIMAL(10, 6) NULL;
  `);
}

/** SQLite may store JSON array; SQL Server now uses comma-separated text. */
function permissionsOrRolesToCsv(raw: unknown): string {
  const s = String(raw ?? "").trim();
  if (!s) return "";
  if (s.startsWith("[")) {
    try {
      const a = JSON.parse(s);
      if (Array.isArray(a)) return a.map((x) => String(x).trim()).filter(Boolean).join(",");
    } catch {
      /* fall through */
    }
  }
  return s.split(",").map((x) => x.trim()).filter(Boolean).join(",");
}

function stepRoleAt(stepsJson: string | null, index: number): string {
  try {
    const steps = JSON.parse(stepsJson || "[]");
    const s = steps[index];
    return String(s?.approverRole || "").toLowerCase();
  } catch {
    return "";
  }
}

async function main() {
  if (!conn) {
    console.error("Set SQLSERVER_CONNECTION_STRING in .env");
    process.exit(1);
  }
  if (!fs.existsSync(sqlitePath)) {
    console.error("SQLite file not found:", sqlitePath);
    process.exit(1);
  }

  const db = new Database(sqlitePath, { readonly: true });
  const pool = await sql.connect(conn);
  console.log("Connected to SQL Server");
  await ensureColumns(pool);

  /** SQLite user id -> SQL Server usersetting.id */
  const userMap = new Map<number, number>();
  const sqliteUsers: { id: number; username: string }[] = db.prepare("SELECT id, username FROM users").all() as any[];
  for (const u of sqliteUsers) {
    const rs = await pool.request().input("username", sql.NVarChar, u.username).query(`SELECT TOP 1 id FROM ${USERS_TABLE} WHERE username = @username`);
    const row = rs.recordset?.[0];
    if (row) userMap.set(u.id, Number(row.id));
    else console.warn("[migrate] No SQL user for SQLite username — rows referencing this id may fail:", u.username, "sqlite id", u.id);
  }

  const count = async (table: string) => {
    const r = await pool.request().query(`SELECT COUNT(*) AS c FROM ${table}`);
    return Number(r.recordset?.[0]?.c ?? 0);
  };

  const existingWorkflows = await count("workflows");
  const existingRequests = await count("workflow_requests");
  if (existingWorkflows > 0 || existingRequests > 0) {
    console.warn(
      `[migrate] Target already has workflows=${existingWorkflows}, workflow_requests=${existingRequests}. ` +
        "Aborting to avoid duplicates. Truncate those tables first if you intend a full re-import."
    );
    await pool.close();
    db.close();
    process.exit(1);
  }

  // --- custom_roles (by name, skip duplicates) ---
  const roles: any[] = db.prepare("SELECT id, name, permissions, created_at FROM custom_roles").all();
  for (const r of roles) {
    const name = String(r.name || "").toLowerCase();
    const perms = permissionsOrRolesToCsv(r.permissions);
    await pool
      .request()
      .input("name", sql.NVarChar, name)
      .input("permissions", sql.NVarChar(sql.MAX), perms)
      .query(
        "IF NOT EXISTS (SELECT 1 FROM custom_roles WHERE name = @name) INSERT INTO custom_roles (name, permissions) VALUES (@name, @permissions)"
      );
  }
  console.log("[migrate] custom_roles merged");

  // --- workflows (preserve ids) ---
  const workflows: any[] = db.prepare("SELECT * FROM workflows").all();
  for (const w of workflows) {
    const creatorSql = userMap.get(Number(w.creator_id));
    if (!creatorSql) {
      console.warn("[migrate] Skip workflow id", w.id, "unknown creator_id", w.creator_id);
      continue;
    }
    await pool
      .request()
      .input("id", sql.Int, w.id)
      .input("creator_id", sql.Int, creatorSql)
      .input("name", sql.NVarChar, w.name)
      .input("category", sql.NVarChar, w.category || "general")
      .input("steps", sql.NVarChar(sql.MAX), w.steps || "[]")
      .input("table_columns", sql.NVarChar(sql.MAX), w.table_columns || "[]")
      .input("attachments_required", sql.Bit, w.attachments_required ? 1 : 0)
      .input("status", sql.NVarChar, w.status || "pending")
      .input("created_at", sql.DateTime2, w.created_at ? new Date(w.created_at) : new Date())
      .query(`
        SET IDENTITY_INSERT workflows ON;
        INSERT INTO workflows (id, creator_id, name, category, steps, table_columns, attachments_required, status, created_at)
        VALUES (@id, @creator_id, @name, @category, @steps, @table_columns, @attachments_required, @status, @created_at);
        SET IDENTITY_INSERT workflows OFF;
      `);
  }
  console.log("[migrate] workflows:", workflows.length);

  // --- attachments (template) ---
  const attW: any[] = db.prepare("SELECT * FROM attachments").all();
  for (const a of attW) {
    await pool
      .request()
      .input("id", sql.Int, a.id)
      .input("workflow_id", sql.Int, a.workflow_id)
      .input("file_name", sql.NVarChar, a.file_name)
      .input("file_type", sql.NVarChar, a.file_type || "")
      .input("file_data", sql.NVarChar(sql.MAX), a.file_data || "")
      .query(`
        SET IDENTITY_INSERT attachments ON;
        INSERT INTO attachments (id, workflow_id, file_name, file_type, file_data) VALUES (@id, @workflow_id, @file_name, @file_type, @file_data);
        SET IDENTITY_INSERT attachments OFF;
      `);
  }
  console.log("[migrate] attachments:", attW.length);

  // --- workflow_requests ---
  const requests: any[] = db.prepare("SELECT * FROM workflow_requests").all();
  const sqliteUserById = new Map(sqliteUsers.map((u) => [u.id, u.username]));
  for (const r of requests) {
    const reqSql = userMap.get(Number(r.requester_id));
    if (!reqSql) {
      console.warn("[migrate] Skip request id", r.id, "unknown requester", r.requester_id);
      continue;
    }
    const requesterName =
      r.requester_username_snapshot ||
      r.requester_name ||
      sqliteUserById.get(Number(r.requester_id)) ||
      "";

    await pool
      .request()
      .input("id", sql.Int, r.id)
      .input("template_id", sql.Int, r.template_id)
      .input("requester_id", sql.Int, reqSql)
      .input("department", sql.NVarChar, r.department || "")
      .input("title", sql.NVarChar, r.title || "")
      .input("details", sql.NVarChar, r.details || "")
      .input("line_items", sql.NVarChar(sql.MAX), r.line_items || "[]")
      .input("entity", sql.NVarChar, r.entity || "")
      .input("formatted_id", sql.NVarChar, r.formatted_id || null)
      .input("status", sql.NVarChar, r.status || "pending")
      .input("current_step_index", sql.Int, r.current_step_index ?? 0)
      .input("requester_signature", sql.NVarChar(sql.MAX), r.requester_signature ?? null)
      .input("requester_signed_at", sql.DateTime2, r.requester_signature ? new Date(r.created_at) : null)
      .input("tax_rate", sql.Decimal(10, 6), r.tax_rate ?? 0.18)
      .input("discount_rate", sql.Decimal(10, 6), (r as any).discount_rate ?? 0)
      .input("currency", sql.NVarChar, String(r.currency ?? "").trim())
      .input("cost_center", sql.NVarChar, r.cost_center || "")
      .input("request_steps", sql.NVarChar(sql.MAX), r.request_steps ?? null)
      .input("created_at", sql.DateTime2, r.created_at ? new Date(r.created_at) : new Date())
      .input("requester_username_snapshot", sql.NVarChar, r.requester_username_snapshot || requesterName)
      .input("template_name_snapshot", sql.NVarChar, r.template_name_snapshot ?? null)
      .input("requester_name", sql.NVarChar, requesterName)
      .query(`
        SET IDENTITY_INSERT workflow_requests ON;
        INSERT INTO workflow_requests (
          id, template_id, requester_id, department, title, details, line_items, entity, formatted_id,
          status, current_step_index, requester_signature, requester_signed_at, tax_rate, discount_rate, currency, cost_center, request_steps, created_at,
          requester_username_snapshot, template_name_snapshot, requester_name
        ) VALUES (
          @id, @template_id, @requester_id, @department, @title, @details, @line_items, @entity, @formatted_id,
          @status, @current_step_index, @requester_signature, @requester_signed_at, @tax_rate, @discount_rate, @currency, @cost_center, @request_steps, @created_at,
          @requester_username_snapshot, @template_name_snapshot, @requester_name
        );
        SET IDENTITY_INSERT workflow_requests OFF;
      `);
  }
  console.log("[migrate] workflow_requests:", requests.length);

  // --- request_attachments ---
  const ra: any[] = db.prepare("SELECT * FROM request_attachments").all();
  for (const x of ra) {
    await pool
      .request()
      .input("id", sql.Int, x.id)
      .input("request_id", sql.Int, x.request_id)
      .input("file_name", sql.NVarChar, x.file_name)
      .input("file_type", sql.NVarChar, x.file_type || "")
      .input("file_data", sql.NVarChar(sql.MAX), x.file_data ?? null)
      .query(`
        SET IDENTITY_INSERT request_attachments ON;
        INSERT INTO request_attachments (id, request_id, file_name, file_type, file_data, file_path) VALUES (@id, @request_id, @file_name, @file_type, @file_data, NULL);
        SET IDENTITY_INSERT request_attachments OFF;
      `);
  }
  console.log("[migrate] request_attachments:", ra.length);

  // --- request_approvals ---
  const appr: any[] = db.prepare("SELECT * FROM request_approvals").all();
  for (const a of appr) {
    const approverSql = userMap.get(Number(a.approver_id));
    if (!approverSql) {
      console.warn("[migrate] Skip approval id", a.id, "unknown approver", a.approver_id);
      continue;
    }
    const uname = sqliteUserById.get(Number(a.approver_id)) || "";
    await pool
      .request()
      .input("id", sql.Int, a.id)
      .input("request_id", sql.Int, a.request_id)
      .input("step_index", sql.Int, a.step_index)
      .input("approver_id", sql.Int, approverSql)
      .input("status", sql.NVarChar, a.status)
      .input("comment", sql.NVarChar, a.comment || "")
      .input("created_at", sql.DateTime2, a.created_at ? new Date(a.created_at) : new Date())
      .input("approver_signature", sql.NVarChar(sql.MAX), a.approver_signature ?? null)
      .input("approver_username", sql.NVarChar, uname)
      .query(`
        SET IDENTITY_INSERT request_approvals ON;
        INSERT INTO request_approvals (
          id, request_id, step_index, approver_id, status, comment, created_at, approver_signature,
          approver_username, approver_role_snapshot, request_title_snapshot, request_formatted_id_snapshot
        ) VALUES (
          @id, @request_id, @step_index, @approver_id, @status, @comment, @created_at, @approver_signature,
          @approver_username, NULL, NULL, NULL
        );
        SET IDENTITY_INSERT request_approvals OFF;
      `);
  }
  console.log("[migrate] request_approvals:", appr.length);

  // --- Backfill checker / final approver on workflow_requests from SQLite approvals + steps ---
  for (const r of requests) {
    const reqSql = userMap.get(Number(r.requester_id));
    if (!reqSql) continue;
    const stepsJson = r.request_steps || (() => {
      const w: any = db.prepare("SELECT steps FROM workflows WHERE id = ?").get(r.template_id);
      return w?.steps || "[]";
    })();
    const approvals: any[] = db
      .prepare("SELECT * FROM request_approvals WHERE request_id = ? ORDER BY created_at ASC")
      .all(r.id);
    let checkedAt: Date | null = null;
    let checkerName: string | null = null;
    let approvedAt: Date | null = null;
    let approverName: string | null = null;
    const stepsArr = JSON.parse(stepsJson || "[]");
    const nSteps = Array.isArray(stepsArr) ? stepsArr.length : 0;
    for (const a of approvals) {
      if (String(a.status).toLowerCase() !== "approved") continue;
      const role = stepRoleAt(stepsJson, Number(a.step_index));
      const uname = sqliteUserById.get(Number(a.approver_id)) || "";
      if (role === "checker") {
        const t = new Date(a.created_at);
        if (!checkedAt || t < checkedAt) {
          checkedAt = t;
          checkerName = uname;
        }
      }
    }
    const approvedApprovals = approvals.filter((x) => String(x.status).toLowerCase() === "approved");
    if (approvedApprovals.length > 0 && nSteps > 0) {
      const last = approvedApprovals[approvedApprovals.length - 1];
      const afterStep = Number(last.step_index) + 1;
      if (afterStep >= nSteps && String(r.status).toLowerCase() === "approved") {
        approvedAt = new Date(last.created_at);
        approverName = sqliteUserById.get(Number(last.approver_id)) || "";
      }
    }
    await pool
      .request()
      .input("id", sql.Int, r.id)
      .input("checked_at", sql.DateTime2, checkedAt)
      .input("checker_name", sql.NVarChar, checkerName)
      .input("approved_at", sql.DateTime2, approvedAt)
      .input("approver_name", sql.NVarChar, approverName)
      .query(`
        UPDATE workflow_requests SET
          checked_at = @checked_at,
          checker_name = @checker_name,
          approved_at = @approved_at,
          approver_name = @approver_name
        WHERE id = @id
      `);
  }
  console.log("[migrate] Backfilled checker/approver audit fields on workflow_requests");

  await pool.close();
  db.close();
  console.log("Done. You can remove better-sqlite3 / database.db after verifying the app.");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
