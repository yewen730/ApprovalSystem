import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import fs from "node:fs";
import os from "node:os";
import sql from "mssql";
import { approvalEventLog, approvalLogSanitize, formatActor } from "./approvalLog";
import {
  COMPANY_FILE_STORAGE_ROOT,
  copyStoredFileToRequest,
  decodeAttachmentPayload,
  getAttachmentsRoot,
  resolveStoredPath,
  saveRequestAttachmentFile,
  tryUnlinkStoredFile,
} from "./attachmentStorage";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import { fileURLToPath } from "url";
import "dotenv/config";

function formatSqlError(e: any): string {
  const msg = String(e?.message || e || "");
  const preceding = Array.isArray(e?.precedingErrors)
    ? e.precedingErrors.map((p: any) => String(p?.message || p || "").trim()).filter(Boolean)
    : [];
  const detail = preceding.length ? ` | precedingErrors: ${preceding.join(" | ")}` : "";
  return `${msg}${detail}`.trim();
}

const SIGNATURE_STORAGE_PATH =
  process.env.SIGNATURE_STORAGE_PATH ||
  "\\\\10.128.3.10\\data\\E_IVOICING\\Approval System\\Signature";

function sanitizeFileComponent(s: string, maxLen = 80): string {
  const raw = String(s || "").trim();
  const safe = raw.replace(/[^\w.\-]+/g, "_").replace(/^_+|_+$/g, "");
  if (!safe) return "user";
  return safe.length > maxLen ? safe.slice(0, maxLen) : safe;
}

function signatureFilePathForUser(user: any): string {
  const id = Number(user?.id);
  const uname = sanitizeFileComponent(user?.username || "user");
  const file = Number.isFinite(id) && id > 0 ? `${id}-${uname}.png` : `${uname}.png`;
  return path.join(SIGNATURE_STORAGE_PATH, file);
}

function parsePngDataUrl(dataUrl: string): Buffer {
  const s = String(dataUrl || "").trim();
  if (!/^data:image\/png;base64,/i.test(s)) {
    throw new Error("Signature must be a PNG data URL");
  }
  const b64 = s.split(",", 2)[1] || "";
  if (!b64) throw new Error("Signature payload is empty");
  return Buffer.from(b64, "base64");
}

function getLanIPv4Addresses(): string[] {
  const nets = os.networkInterfaces();
  const out = new Set<string>();
  const isV4 = (family: string | number) => family === "IPv4" || family === 4;
  for (const list of Object.values(nets)) {
    if (!list) continue;
    for (const net of list) {
      if (!isV4(net.family) || net.internal) continue;
      if (net.address.startsWith("169.254.")) continue;
      out.add(net.address);
    }
  }
  return [...out];
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const JWT_SECRET = process.env.JWT_SECRET || "";
if (!JWT_SECRET) {
  throw new Error("Missing JWT_SECRET in environment (.env)");
}
const SQLSERVER_CONNECTION_STRING = process.env.SQLSERVER_CONNECTION_STRING || "";
let sqlPool: sql.ConnectionPool;
const rolePermissionsCache = new Map<string, string[]>();

/** SQL Server auth + per-user approval limit (column `approval_limit`, MYR). Default: [dbo].[usersetting] */
const DEFAULT_SQL_USERS_TABLE = "[dbo].[usersetting]";
const rawSqlUsersTable = (process.env.SQLSERVER_USERS_TABLE || DEFAULT_SQL_USERS_TABLE).trim();
const SQLSERVER_USERS_TABLE =
  /^[\w\[\].]+$/.test(rawSqlUsersTable) && rawSqlUsersTable.length < 200 ? rawSqlUsersTable : DEFAULT_SQL_USERS_TABLE;

/** PR: single approver step (must match client FIXED_PR_STEPS). */
const FIXED_PR_STEPS = [
  { id: "pr-step-1", label: "Approver", approverRole: "approver" },
];

/** PO template: full chain; director step may be omitted per request when total ≤ RM30k equivalent. */
const FIXED_PO_STEPS_FULL = [
  { id: "po-step-1", label: "Checker Verification", approverRole: "checker" },
  { id: "po-step-2", label: "Final Approval", approverRole: "som" },
  { id: "po-step-3", label: "Director Authorization (> RM100,000)", approverRole: "director" },
];

const FIXED_PO_STEPS_BASE = FIXED_PO_STEPS_FULL.slice(0, 2);

/** SR: exactly two approval steps (must match requested flow). */
const FIXED_SR_STEPS = [
  { id: "sr-step-1", label: "HOD Approval", approverRole: "approver" },
  { id: "sr-step-2", label: "Final Approval", approverRole: "som" },
];

/** Total PO/PR value above this (in MYR equivalent) requires director step on PO. */
const PO_DIRECTOR_THRESHOLD_MYR = 100000;

/** Rough FX to MYR for threshold checks (tune to business rates). */
const MYR_PER_UNIT: Record<string, number> = {
  MYR: 1,
  RM: 1,
  USD: 4.47,
  SGD: 3.55,
  EUR: 5.0,
  EURO: 5.0,
  GBP: 5.65,
  FCFA: 0.0075,
};

/** Stock / spare requisition (not purchase request "PR"). */
const isSRName = (name: string) => {
  const n = name.toLowerCase();
  return n.includes("stock requisition") || (n.includes("stock") && n.includes("requisition"));
};

const isPRName = (name: string) => {
  const n = name.toLowerCase();
  if (isSRName(name)) return false;
  return n.includes("purchase request") || n.includes("pr");
};

const isPOName = (name: string) => {
  const n = name.toLowerCase();
  return n.includes("purchase order") || n.includes("po");
};

function getLineItemValue(item: any, keys: string[]): string {
  const foundKey = Object.keys(item || {}).find((k) => keys.includes(k.toLowerCase()));
  if (!foundKey) return "";
  return String(item[foundKey] ?? "");
}

function convertToMyr(amount: number, currency: string): number {
  const c = (currency ?? "").trim().toUpperCase();
  // No stored currency: treat face value as MYR for threshold math only (not persisted as USD).
  const rate = c === "" ? 1 : MYR_PER_UNIT[c] ?? MYR_PER_UNIT.USD;
  return amount * rate;
}

function clampUnitRate(n: number): number {
  if (Number.isNaN(n) || n < 0) return 0;
  if (n > 1) return 1;
  return n;
}

/** Grand total in MYR after optional discount on subtotal, then tax on the discounted base. */
function lineItemQtyForTotals(item: any): number {
  const raw = item || {};
  const lowerKeys = new Map(Object.keys(raw).map((k) => [k.toLowerCase(), k]));
  const pickNum = (...names: string[]) => {
    for (const n of names) {
      const k = lowerKeys.get(n);
      if (k != null) {
        const v = parseFloat(String(raw[k] ?? ""));
        if (!Number.isNaN(v) && v > 0) return v;
      }
    }
    return 0;
  };
  const maxQ = pickNum("max quantity");
  if (maxQ > 0) return maxQ;
  const qty = pickNum("quantity", "qty");
  if (qty > 0) return qty;
  const minQ = pickNum("min quantity");
  return Number.isFinite(minQ) ? minQ : 0;
}

function computeRequestTotalMyr(
  lineItems: any[],
  taxRate: number | undefined,
  currency: string,
  discountRate?: number | null
): number {
  let subtotal = 0;
  for (const item of lineItems || []) {
    const qty = lineItemQtyForTotals(item);
    const price = parseFloat(getLineItemValue(item, ["unit price", "price", "amount"]) || "0");
    subtotal += qty * price;
  }
  const tr = clampUnitRate(taxRate !== undefined ? Number(taxRate) : 0.18);
  const dr = clampUnitRate(discountRate !== undefined && discountRate !== null ? Number(discountRate) : 0);
  const afterDiscount = subtotal * (1 - dr);
  const total = afterDiscount + afterDiscount * tr;
  return convertToMyr(total, currency);
}

function buildPoRequestSteps(totalMyr: number) {
  const steps = [...FIXED_PO_STEPS_BASE];
  if (totalMyr > PO_DIRECTOR_THRESHOLD_MYR) {
    steps.push(FIXED_PO_STEPS_FULL[2]);
  }
  return steps;
}

const seedRoles = [
  { name: "admin", permissions: ["admin", "view_history", "create_templates", "approve_templates", "manage_users", "edit_requests", "view_procurement_center"] },
  { name: "preparer", permissions: [] },
  { name: "checker", permissions: ["create_templates", "approve_templates", "edit_requests"] },
  { name: "approver", permissions: ["create_templates", "approve_templates", "edit_requests"] },
  { name: "director", permissions: ["view_history", "approve_templates"] },
  { name: "som", permissions: ["view_history", "approve_templates", "view_procurement_center", "edit_requests"] },
  { name: "purchasing", permissions: ["view_procurement_center"] },
  { name: "user", permissions: [] },
];

async function normalizeProcurementWorkflowSteps(pool: sql.ConnectionPool) {
  const rs = await pool.request().query("SELECT id, name FROM workflows WHERE category = N'procurement'");
  for (const row of rs.recordset || []) {
    const w: any = row;
    if (isPOName(w.name)) {
      await pool
        .request()
        .input("steps", sql.NVarChar(sql.MAX), JSON.stringify(FIXED_PO_STEPS_FULL))
        .input("id", sql.Int, w.id)
        .query("UPDATE workflows SET steps = @steps WHERE id = @id");
    } else if (isPRName(w.name)) {
      await pool
        .request()
        .input("steps", sql.NVarChar(sql.MAX), JSON.stringify(FIXED_PR_STEPS))
        .input("id", sql.Int, w.id)
        .query("UPDATE workflows SET steps = @steps WHERE id = @id");
    } else if (isSRName(w.name)) {
      await pool
        .request()
        .input("steps", sql.NVarChar(sql.MAX), JSON.stringify(FIXED_SR_STEPS))
        .input("id", sql.Int, w.id)
        .query("UPDATE workflows SET steps = @steps WHERE id = @id");
    }
  }
}

/**
 * Simple token lists in the DB (`usersetting.role`, `usersetting.entities`, `custom_roles.permissions`):
 * comma-separated text. Still reads legacy JSON arrays `["a","b"]` for old rows.
 * Large / nested payloads (workflow steps, line_items, table_columns, etc.) stay JSON in NVARCHAR.
 */
const parseCommaSeparatedList = (value: any): string[] => {
  if (Array.isArray(value)) return value.map((v) => String(v).trim()).filter(Boolean);
  if (value == null || value === "") return [];
  const s = String(value).trim();
  if (!s) return [];
  if (s.startsWith("[")) {
    try {
      const parsed = JSON.parse(s);
      if (Array.isArray(parsed)) return parsed.map((v) => String(v).trim()).filter(Boolean);
    } catch {
      /* fall through */
    }
  }
  return s.split(",").map((x) => x.trim()).filter(Boolean);
};

const toCommaSeparated = (items: string[] | undefined | null): string =>
  (items || []).map((x) => String(x).trim()).filter(Boolean).join(",");

/** Normalize API input (array, CSV string, or legacy JSON) to DB CSV text. */
const listToStoredCsv = (value: any): string => toCommaSeparated(parseCommaSeparatedList(value));

const parseNullableNumber = (value: any): number | null => {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
};

const userHasDepartment = (departmentCsv: any, expectedDepartment: string): boolean =>
  parseCommaSeparatedList(departmentCsv).some(
    (d) => d.toLowerCase() === String(expectedDepartment || "").toLowerCase()
  );

/** Plain text or bcrypt (`$2a$` / `$2b$`) passwords in SQL usersetting. */
const sqlPasswordMatches = (stored: string | null | undefined, plain: string): boolean => {
  const s = String(stored ?? "");
  if (s.startsWith("$2a$") || s.startsWith("$2b$") || s.startsWith("$2y$")) {
    try {
      return bcrypt.compareSync(plain, s);
    } catch {
      return false;
    }
  }
  return s === String(plain ?? "");
};

/**
 * Approver cap in MYR for this entity: `entity_approver_registry` first (explicit amount),
 * then legacy `user_entity_approval_limits`, then global `usersetting.approval_limit`.
 * GCCM approvals skip this at the HTTP layer (see approve handler).
 */
const getUserApprovalLimitMyr = async (
  pool: sql.ConnectionPool,
  userId: number,
  entityCode?: string | null
): Promise<number | null> => {
  if (!userId) return null;
  const id = Number(userId);
  const ent = (entityCode ?? "").toString().trim().toUpperCase();
  if (ent) {
    try {
      const reg = await pool
        .request()
        .input("userId", sql.Int, id)
        .input("entity", sql.NVarChar, ent)
        .query(
          `SELECT TOP 1 approval_limit_myr AS lim FROM dbo.entity_approver_registry WHERE user_id = @userId AND UPPER(LTRIM(RTRIM(entity))) = @entity AND active = 1`
        );
      if (reg.recordset?.length) {
        const lim = parseNullableNumber((reg.recordset[0] as any)?.lim);
        if (lim !== null) return lim;
      }
    } catch {
      /* registry table may not exist yet */
    }
    try {
      const er = await pool
        .request()
        .input("userId", sql.Int, id)
        .input("entity", sql.NVarChar, ent)
        .query(
          `SELECT TOP 1 approval_limit_myr AS lim FROM dbo.user_entity_approval_limits WHERE user_id = @userId AND UPPER(LTRIM(RTRIM(entity))) = @entity`
        );
      if (er.recordset?.length) {
        return parseNullableNumber((er.recordset[0] as any)?.lim);
      }
    } catch {
      /* table may not exist yet */
    }
  }
  const tryQueries = [
    `SELECT TOP 1 approval_limit AS lim FROM ${SQLSERVER_USERS_TABLE} WHERE id = @userId`,
    `SELECT TOP 1 approval_limit_myr AS lim FROM ${SQLSERVER_USERS_TABLE} WHERE id = @userId`,
  ];
  for (const q of tryQueries) {
    try {
      const rs = await pool.request().input("userId", sql.Int, id).query(q);
      const lim = parseNullableNumber(rs.recordset?.[0]?.lim);
      if (lim !== null) return lim;
    } catch {
      // Column may not exist yet
    }
  }
  return null;
};

const GCCM_ENTITY = "GCCM";

function stepsJsonIncludesApproverRole(stepsJson: string | null | undefined): boolean {
  try {
    const arr = JSON.parse(String(stepsJson || "[]"));
    if (!Array.isArray(arr)) return false;
    return arr.some((s: any) => String(s?.approverRole || "").toLowerCase() === "approver");
  } catch {
    return false;
  }
}

function userRowHasEntityAccess(row: any, entityUpper: string): boolean {
  return parseCommaSeparatedList(row?.entities).some((e) => e.toUpperCase() === entityUpper.toUpperCase());
}

function userRowHasApproverRole(row: any): boolean {
  return parseCommaSeparatedList(row?.role).some((r) => r.toLowerCase() === "approver");
}

/** Same department string rule as workflow request rows vs JWT user (see assertWorkflowRequestAccess). */
function userDepartmentsMatch(a: string | null | undefined, b: string | null | undefined): boolean {
  const left = new Set(
    parseCommaSeparatedList(a)
      .map((x) => x.toLowerCase())
      .filter(Boolean)
  );
  const right = parseCommaSeparatedList(b)
    .map((x) => x.toLowerCase())
    .filter(Boolean);
  if (left.size === 0 || right.length === 0) return false;
  return right.some((d) => left.has(d));
}

async function fetchSqlUsersettingById(pool: sql.ConnectionPool, userId: number): Promise<any | null> {
  try {
    const rs = await pool.request().input("id", sql.Int, Number(userId)).query(`SELECT TOP 1 * FROM ${SQLSERVER_USERS_TABLE} WHERE id = @id`);
    return rs.recordset?.[0] ?? null;
  } catch {
    return null;
  }
}

/** Load one user row from SQL Server `usersetting`. */
const fetchSqlUsersettingByUsername = async (pool: sql.ConnectionPool, username: string): Promise<any | null> => {
  const base = `SELECT TOP 1 id, username, password, role, department, entities FROM ${SQLSERVER_USERS_TABLE} WHERE username = @username`;
  const withLimit = `SELECT TOP 1 id, username, password, role, department, entities, approval_limit FROM ${SQLSERVER_USERS_TABLE} WHERE username = @username`;
  try {
    const r = await pool.request().input("username", sql.NVarChar, username).query(withLimit);
    if (r.recordset?.[0]) return r.recordset[0];
  } catch {
    // `approval_limit` column may not exist yet
  }
  try {
    const r2 = await pool.request().input("username", sql.NVarChar, username).query(base);
    return r2.recordset?.[0] ?? null;
  } catch (e: any) {
    console.warn("SQL usersetting lookup failed:", e?.message || e);
    return null;
  }
};

const refreshRolePermissionsCache = async (pool: sql.ConnectionPool) => {
  rolePermissionsCache.clear();
  const rs = await pool.request().query("SELECT name, permissions FROM custom_roles");
  for (const r of rs.recordset || []) {
    const name = String((r as any).name || "").toLowerCase();
    const perms = parseCommaSeparatedList((r as any).permissions);
    rolePermissionsCache.set(name, perms);
  }
};

const quoteSqlServerIdent = (name: string): string => `[${String(name || "").replace(/]/g, "]]")}]`;

type CostCenterTableInfo = {
  schemaName: string;
  tableName: string;
  qualifiedName: string;
  hasId: boolean;
};

/** Resolve the physical SQL table used for cost-center catalogs (strictly `cost_center`). */
async function resolveCostCenterTableInfo(pool: sql.ConnectionPool): Promise<CostCenterTableInfo> {
  const rs = await pool.request().query(`
    SELECT TOP 1
      s.name AS schema_name,
      t.name AS table_name,
      CASE WHEN EXISTS (
        SELECT 1 FROM sys.columns c WHERE c.object_id = t.object_id AND c.name = N'id'
      ) THEN 1 ELSE 0 END AS has_id
    FROM sys.tables t
    JOIN sys.schemas s ON s.schema_id = t.schema_id
    WHERE t.name = N'cost_center'
    ORDER BY CASE WHEN s.name = N'dbo' THEN 0 ELSE 1 END, s.name
  `);
  const row = rs.recordset?.[0] as { schema_name?: unknown; table_name?: unknown; has_id?: unknown } | undefined;
  const schemaName = String(row?.schema_name || "").trim();
  const tableName = String(row?.table_name || "").trim();
  if (!schemaName || !tableName) {
    throw new Error("Required table `cost_center` was not found in this SQL Server database.");
  }
  return {
    schemaName,
    tableName,
    qualifiedName: `${quoteSqlServerIdent(schemaName)}.${quoteSqlServerIdent(tableName)}`,
    hasId: Number(row?.has_id ?? 0) === 1,
  };
}

/**
 * Ensures `cost_center` exists and has columns used by GET /api/cost-centers.
 * Called at startup and again on each cost-centers read so a failed startup migration can self-heal.
 */
async function ensureCostCentersTable(pool: sql.ConnectionPool): Promise<void> {
  await pool.request().query(`
    IF NOT EXISTS (SELECT 1 FROM sys.tables t WHERE t.name = N'cost_center')
    BEGIN
      CREATE TABLE dbo.cost_center (
        id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_cost_centers PRIMARY KEY,
        entity NVARCHAR(32) NOT NULL,
        code NVARCHAR(64) NOT NULL,
        name NVARCHAR(500) NOT NULL,
        gl_account NVARCHAR(128) NULL,
        approval NVARCHAR(512) NULL,
        status BIT NOT NULL CONSTRAINT DF_cost_centers_status DEFAULT (1),
        created_at DATETIME2(3) NOT NULL CONSTRAINT DF_cost_centers_created DEFAULT SYSUTCDATETIME(),
        CONSTRAINT UQ_cost_centers_entity_code UNIQUE (entity, code)
      );
      CREATE INDEX IX_cost_centers_entity_status_code ON dbo.cost_center (entity, status, code);
    END
  `);
  const tableInfo = await resolveCostCenterTableInfo(pool);
  const colLengthTarget = `${tableInfo.schemaName}.${tableInfo.tableName}`.replace(/'/g, "''");
  await pool.request().query(`
    IF OBJECT_ID(N'${colLengthTarget}', N'U') IS NOT NULL
    BEGIN
      IF COL_LENGTH('${colLengthTarget}', 'entity') IS NULL
        ALTER TABLE ${tableInfo.qualifiedName} ADD entity NVARCHAR(32) NULL;
      IF COL_LENGTH('${colLengthTarget}', 'gl_account') IS NULL
        ALTER TABLE ${tableInfo.qualifiedName} ADD gl_account NVARCHAR(128) NULL;
      IF COL_LENGTH('${colLengthTarget}', 'approval') IS NULL
        ALTER TABLE ${tableInfo.qualifiedName} ADD approval NVARCHAR(512) NULL;
      IF COL_LENGTH('${colLengthTarget}', 'status') IS NULL
        ALTER TABLE ${tableInfo.qualifiedName} ADD status BIT NOT NULL CONSTRAINT DF_cc_status_patch DEFAULT (1);
      IF COL_LENGTH('${colLengthTarget}', 'created_at') IS NULL
        ALTER TABLE ${tableInfo.qualifiedName} ADD created_at DATETIME2(3) NOT NULL CONSTRAINT DF_cc_created_patch DEFAULT SYSUTCDATETIME();
      IF EXISTS (
        SELECT 1
        FROM sys.key_constraints kc
        WHERE kc.parent_object_id = OBJECT_ID(N'${colLengthTarget}')
          AND kc.name = N'UQ_cost_centers_code'
      )
      BEGIN
        ALTER TABLE ${tableInfo.qualifiedName} DROP CONSTRAINT UQ_cost_centers_code;
      END
    END
  `);

  // If the table existed before this app version, it may contain duplicates that prevent
  // adding the (entity, code) unique constraint. We dedupe safely by keeping the lowest id.
  if (tableInfo.hasId) {
    try {
      await pool.request().query(`
        WITH d AS (
          SELECT
            id,
            ROW_NUMBER() OVER (
              PARTITION BY UPPER(LTRIM(RTRIM(entity))), UPPER(LTRIM(RTRIM(code)))
              ORDER BY id ASC
            ) AS rn
          FROM ${tableInfo.qualifiedName}
          WHERE entity IS NOT NULL AND LTRIM(RTRIM(entity)) <> N'' AND code IS NOT NULL AND LTRIM(RTRIM(code)) <> N''
        )
        DELETE FROM d WHERE rn > 1;
      `);
    } catch (e: any) {
      console.warn("SQL cost_centers dedupe skipped:", formatSqlError(e));
    }
  }

  // Add entity+code uniqueness if possible; if it fails, keep serving the catalog.
  try {
    await pool.request().query(`
      IF NOT EXISTS (
        SELECT 1
        FROM sys.key_constraints kc
        WHERE kc.parent_object_id = OBJECT_ID(N'${colLengthTarget}')
          AND kc.name = N'UQ_cost_centers_entity_code'
      )
      BEGIN
        ALTER TABLE ${tableInfo.qualifiedName}
        ADD CONSTRAINT UQ_cost_centers_entity_code UNIQUE (entity, code);
      END
    `);
  } catch (e: any) {
    console.warn("SQL cost_centers unique(entity,code) skipped:", formatSqlError(e));
  }

  await pool.request().query(`
    IF NOT EXISTS (
      SELECT 1
      FROM sys.indexes i
      WHERE i.object_id = OBJECT_ID(N'${colLengthTarget}')
        AND i.name = N'IX_cost_centers_entity_status_code'
    )
    BEGIN
      CREATE INDEX IX_cost_centers_entity_status_code ON ${tableInfo.qualifiedName} (entity, status, code);
    END
  `);
  const cntRs = await pool.request().query(`SELECT COUNT(*) AS c FROM ${tableInfo.qualifiedName}`);
  const n = Number((cntRs.recordset?.[0] as any)?.c ?? 0);
  if (n === 0) {
    await pool.request().query(`
      INSERT INTO ${tableInfo.qualifiedName} (entity, code, name, gl_account, approval, status) VALUES
      (N'GCCM', N'1000', N'IT Department', NULL, NULL, 1),
      (N'GCCM', N'2000', N'HR Department', NULL, NULL, 1),
      (N'GCCM', N'3000', N'Finance Department', NULL, NULL, 1),
      (N'GCCM', N'4000', N'Marketing', NULL, NULL, 1),
      (N'GCCM', N'5000', N'Operations', NULL, NULL, 1),
      (N'GCCM', N'6000', N'Sales', NULL, NULL, 1),
      (N'GCCM', N'7000', N'R&D', NULL, NULL, 1),
      (N'GCCM', N'8000', N'Legal', NULL, NULL, 1),
      (N'GCCM', N'9000', N'Administration', NULL, NULL, 1);
    `);
  }
}

/** Optional columns for workflow / request tables on SQL Server (safe no-ops if already present). */
const ensureSqlServerWorkflowColumns = async (pool: sql.ConnectionPool) => {
  try {
    await pool.request().query(`
      IF COL_LENGTH('request_attachments', 'file_path') IS NULL
        ALTER TABLE request_attachments ADD file_path NVARCHAR(1000) NULL;
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
      IF COL_LENGTH('workflow_requests', 'section') IS NULL
        ALTER TABLE workflow_requests ADD section NVARCHAR(255) NULL;
      IF COL_LENGTH('workflow_requests', 'suggested_supplier') IS NULL
        ALTER TABLE workflow_requests ADD suggested_supplier NVARCHAR(512) NULL;
      IF COL_LENGTH('workflow_requests', 'converted_po_request_id') IS NULL
        ALTER TABLE workflow_requests ADD converted_po_request_id INT NULL;
    `);
  } catch (e: any) {
    console.warn("SQL workflow column ensure skipped:", e?.message || e);
  }
  /** Base64 PDFs exceed typical NVARCHAR(4000) schemas; widen so inserts match sql.NVarChar(sql.MAX). */
  const widenAttachmentFileData = [
    `IF COL_LENGTH('request_attachments', 'file_data') IS NOT NULL
      ALTER TABLE request_attachments ALTER COLUMN file_data NVARCHAR(MAX) NULL;`,
  ];
  for (const q of widenAttachmentFileData) {
    try {
      await pool.request().query(q);
    } catch (e: any) {
      console.warn("SQL widen file_data to NVARCHAR(MAX):", e?.message || e);
    }
  }
  const widenSignatureColumns = [
    `IF COL_LENGTH('workflow_requests', 'requester_signature') IS NOT NULL
      ALTER TABLE workflow_requests ALTER COLUMN requester_signature NVARCHAR(MAX) NULL;`,
    `IF COL_LENGTH('request_approvals', 'approver_signature') IS NOT NULL
      ALTER TABLE request_approvals ALTER COLUMN approver_signature NVARCHAR(MAX) NULL;`,
  ];
  for (const q of widenSignatureColumns) {
    try {
      await pool.request().query(q);
    } catch (e: any) {
      console.warn("SQL widen signature columns to NVARCHAR(MAX):", e?.message || e);
    }
  }
  try {
    await pool.request().query(`
      IF COL_LENGTH('workflow_requests', 'assigned_approver_id') IS NULL
        ALTER TABLE workflow_requests ADD assigned_approver_id INT NULL;
    `);
  } catch (e: any) {
    console.warn("SQL assigned_approver_id column:", e?.message || e);
  }
  try {
    await pool.request().query(`
      IF COL_LENGTH('workflows', 'is_active') IS NULL
        ALTER TABLE workflows ADD is_active BIT NOT NULL CONSTRAINT DF_workflows_is_active DEFAULT (1);
    `);
    await pool.request().query(`
      IF COL_LENGTH('workflows', 'is_active') IS NOT NULL
        UPDATE workflows SET is_active = 1 WHERE is_active IS NULL;
    `);
  } catch (e: any) {
    console.warn("SQL workflows.is_active column:", e?.message || e);
  }
  try {
    const utEsc = SQLSERVER_USERS_TABLE.replace(/'/g, "''");
    await pool.request().query(`
      IF NOT EXISTS (
        SELECT 1 FROM sys.columns c
        WHERE c.object_id = OBJECT_ID(N'${utEsc}') AND c.name = N'designation'
      )
      BEGIN
        ALTER TABLE ${SQLSERVER_USERS_TABLE} ADD designation NVARCHAR(255) NULL;
      END
    `);
  } catch (e: any) {
    console.warn("SQL usersetting.designation column:", e?.message || e);
  }
  try {
    await pool.request().query(`
      IF OBJECT_ID(N'dbo.user_entity_approval_limits', N'U') IS NULL
      BEGIN
        CREATE TABLE dbo.user_entity_approval_limits (
          user_id INT NOT NULL,
          entity NVARCHAR(32) NOT NULL,
          approval_limit_myr DECIMAL(18, 2) NULL,
          CONSTRAINT PK_user_entity_approval_limits PRIMARY KEY (user_id, entity)
        );
      END
    `);
  } catch (e: any) {
    console.warn("SQL user_entity_approval_limits table:", e?.message || e);
  }
  try {
    await pool.request().query(`
      IF OBJECT_ID(N'dbo.entity_approver_registry', N'U') IS NULL
      BEGIN
        CREATE TABLE dbo.entity_approver_registry (
          entity NVARCHAR(32) NOT NULL,
          user_id INT NOT NULL,
          selectable_by_requestor BIT NOT NULL CONSTRAINT DF_entity_appr_reg_pick DEFAULT (0),
          approval_limit_myr DECIMAL(18, 2) NULL,
          active BIT NOT NULL CONSTRAINT DF_entity_appr_reg_active DEFAULT (1),
          sort_order INT NULL,
          CONSTRAINT PK_entity_approver_registry PRIMARY KEY (entity, user_id),
          CONSTRAINT CK_entity_appr_reg_entity_nonempty CHECK (LEN(LTRIM(RTRIM(entity))) > 0)
        );
        CREATE NONCLUSTERED INDEX IX_entity_appr_reg_entity_active_pick
          ON dbo.entity_approver_registry (entity, active, selectable_by_requestor)
          INCLUDE (user_id, sort_order);
      END
    `);
  } catch (e: any) {
    console.warn("SQL entity_approver_registry table:", e?.message || e);
  }
  try {
    await ensureCostCentersTable(pool);
  } catch (e: any) {
    console.warn("SQL cost_centers table:", formatSqlError(e));
  }
  try {
    await pool.request().query(`
      IF OBJECT_ID(N'dbo.spare_locations', N'U') IS NULL
      BEGIN
        CREATE TABLE dbo.spare_locations (
          id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_spare_locations PRIMARY KEY,
          entity NVARCHAR(32) NOT NULL,
          code NVARCHAR(64) NOT NULL,
          name NVARCHAR(500) NOT NULL,
          status BIT NOT NULL CONSTRAINT DF_spare_locations_status DEFAULT (1),
          created_at DATETIME2(3) NOT NULL CONSTRAINT DF_spare_locations_created DEFAULT SYSUTCDATETIME(),
          CONSTRAINT UQ_spare_locations_entity_code UNIQUE (entity, code),
          CONSTRAINT CK_spare_locations_entity_nonempty CHECK (LEN(LTRIM(RTRIM(entity))) > 0)
        );
        CREATE INDEX IX_spare_locations_entity_status ON dbo.spare_locations (entity, status, code);
      END
    `);
  } catch (e: any) {
    console.warn("SQL spare_locations table:", e?.message || e);
  }
};

const seedSqlRolesIfMissing = async (pool: sql.ConnectionPool) => {
  for (const role of seedRoles) {
    const roleName = String(role.name).toLowerCase();
    const permsCsv = toCommaSeparated((role.permissions || []).map((p: string) => String(p)));
    const existsRs = await pool
      .request()
      .input("name", sql.NVarChar, roleName)
      .query("SELECT TOP 1 id, permissions FROM custom_roles WHERE name = @name");
    const existing: any = existsRs.recordset?.[0];
    if (!existing) {
      await pool
        .request()
        .input("name", sql.NVarChar, roleName)
        .input("permissions", sql.NVarChar(sql.MAX), permsCsv)
        .query("INSERT INTO custom_roles (name, permissions) VALUES (@name, @permissions)");
      continue;
    }
    const existingPerms = parseCommaSeparatedList(existing.permissions);
    if (existingPerms.length === 0) {
      await pool
        .request()
        .input("name", sql.NVarChar, roleName)
        .input("permissions", sql.NVarChar(sql.MAX), permsCsv)
        .query("UPDATE custom_roles SET permissions = @permissions WHERE name = @name");
    }
  }
};

/** If cell still stores legacy `["a","b"]`, rewrite to `a,b` so the DB matches the CSV convention. */
/** True if client completed the requester signature pad (data URL or explicit flag). */
const bodyIndicatesRequesterSigned = (body: any): boolean => {
  if (body?.requester_signed === true || body?.requester_signed === 1) return true;
  const s = body?.requester_signature;
  return typeof s === "string" && s.trim().length > 0;
};

const bodyIndicatesApproverSigned = (body: any): boolean => {
  if (body?.approver_signed === true || body?.approver_signed === 1) return true;
  const s = body?.approver_signature;
  return typeof s === "string" && s.trim().length > 0;
};

const legacyJsonArrayCellToCsv = (raw: unknown): string | null => {
  if (raw == null) return null;
  const s = String(raw).trim();
  if (!s.startsWith("[")) return null;
  try {
    const parsed = JSON.parse(s);
    if (!Array.isArray(parsed)) return null;
    return toCommaSeparated(parsed.map((v) => String(v).trim()).filter(Boolean));
  } catch {
    return null;
  }
};

const normalizeLegacySimpleListJsonToCsv = async (pool: sql.ConnectionPool) => {
  let updatedRoles = 0;
  let updatedUserRoles = 0;
  let updatedUserEntities = 0;
  try {
    const cr = await pool.request().query("SELECT id, permissions FROM custom_roles");
    for (const row of cr.recordset || []) {
      const csv = legacyJsonArrayCellToCsv(row.permissions);
      if (csv === null) continue;
      await pool
        .request()
        .input("id", sql.Int, row.id)
        .input("permissions", sql.NVarChar(sql.MAX), csv)
        .query("UPDATE custom_roles SET permissions = @permissions WHERE id = @id");
      updatedRoles++;
    }
  } catch (e: any) {
    console.warn("normalize custom_roles.permissions:", e?.message || e);
  }
  try {
    const ur = await pool.request().query(`SELECT id, role, entities FROM ${SQLSERVER_USERS_TABLE}`);
    for (const row of ur.recordset || []) {
      const roleCsv = legacyJsonArrayCellToCsv(row.role);
      if (roleCsv !== null) {
        await pool
          .request()
          .input("id", sql.Int, row.id)
          .input("role", sql.NVarChar(sql.MAX), roleCsv)
          .query(`UPDATE ${SQLSERVER_USERS_TABLE} SET role = @role WHERE id = @id`);
        updatedUserRoles++;
      }
      const entCsv = legacyJsonArrayCellToCsv(row.entities);
      if (entCsv !== null) {
        await pool
          .request()
          .input("id", sql.Int, row.id)
          .input("entities", sql.NVarChar(sql.MAX), entCsv)
          .query(`UPDATE ${SQLSERVER_USERS_TABLE} SET entities = @entities WHERE id = @id`);
        updatedUserEntities++;
      }
    }
  } catch (e: any) {
    console.warn("normalize usersetting role/entities:", e?.message || e);
  }
  if (updatedRoles > 0 || updatedUserRoles > 0 || updatedUserEntities > 0) {
    console.log(
      `Normalized legacy JSON → CSV: custom_roles=${updatedRoles}, user role=${updatedUserRoles}, user entities=${updatedUserEntities} (${SQLSERVER_USERS_TABLE})`
    );
  }
};

async function startServer() {
  const conn = (SQLSERVER_CONNECTION_STRING || "").trim();
  if (!conn) {
    console.error("SQLSERVER_CONNECTION_STRING is required (set in .env). Import legacy data: npm run migrate:sqlite");
    process.exit(1);
  }
  try {
    sqlPool = await sql.connect(conn);
  } catch (err: any) {
    console.error("SQL Server connection failed:", err?.message || err);
    process.exit(1);
  }
  console.log("SQL Server connection ready");
  try {
    await ensureSqlServerWorkflowColumns(sqlPool);
  } catch (e: any) {
    console.warn("SQL workflow schema ensure failed:", e?.message || e);
  }
  try {
    await seedSqlRolesIfMissing(sqlPool);
  } catch (seedErr: any) {
    console.warn("SQL role seed failed:", seedErr?.message || seedErr);
  }
  try {
    await normalizeLegacySimpleListJsonToCsv(sqlPool);
  } catch (normErr: any) {
    console.warn("Legacy JSON→CSV normalize:", normErr?.message || normErr);
  }
  try {
    await normalizeProcurementWorkflowSteps(sqlPool);
  } catch (e: any) {
    console.warn("normalizeProcurementWorkflowSteps:", e?.message || e);
  }
  await refreshRolePermissionsCache(sqlPool);

  const app = express();
  app.use(express.json({ limit: "50mb" }));

  console.log("Request attachment storage (ATTACHMENTS_STORAGE_PATH):", getAttachmentsRoot());

  const unlinkAllAttachmentsForRequest = async (requestId: number) => {
    const rs = await sqlPool
      .request()
      .input("request_id", sql.Int, requestId)
      .query("SELECT file_path FROM request_attachments WHERE request_id = @request_id");
    for (const row of rs.recordset || []) {
      tryUnlinkStoredFile((row as any).file_path);
    }
  };

  const PORT = Number(process.env.PORT) || 3000;

  // Helper to get all permissions for a user's roles
  const getUserPermissions = (roleNames: string[]) => {
    if (!roleNames || roleNames.length === 0) return [];
    
    const allPermissions = new Set<string>();
    const lowerRoles = roleNames.map(r => r.toLowerCase());

    if (lowerRoles.includes('admin')) {
      ['admin', 'view_history', 'create_templates', 'approve_templates', 'manage_users', 'edit_requests', 'view_procurement_center'].forEach(p => allPermissions.add(p));
      return Array.from(allPermissions);
    }
    
    // Default permissions for specific roles as requested
    // Approver, Checker can create templates
    if (lowerRoles.includes('approver') || lowerRoles.includes('checker')) {
      allPermissions.add('create_templates');
    }
    // Approver, Checker, Director can approve templates
    if (lowerRoles.includes('approver') || lowerRoles.includes('checker') || lowerRoles.includes('director') || lowerRoles.includes('som')) {
      allPermissions.add('approve_templates');
    }
    
    lowerRoles.forEach((roleName) => {
      (rolePermissionsCache.get(roleName) || []).forEach((p) => allPermissions.add(p));
    });
    
    return Array.from(allPermissions);
  };

  // Middleware for Auth
  const authenticate = (req: any, res: any, next: any) => {
    const token = req.headers.authorization?.split(" ")[1];
    if (!token) return res.status(401).json({ error: "Unauthorized" });
    try {
      const decoded: any = jwt.verify(token, JWT_SECRET);
      // Compatibility for old tokens
      if (!decoded.roles) {
        decoded.roles = decoded.role ? [decoded.role] : ['user'];
      }
      
      // Fetch fresh permissions
      decoded.permissions = getUserPermissions(decoded.roles);
      
      req.user = decoded;
      next();
    } catch (err) {
      res.status(401).json({ error: "Invalid token" });
    }
  };

  const hasPermission = (permission: string) => (req: any, res: any, next: any) => {
    if (!req.user.permissions || (!req.user.permissions.includes(permission) && !req.user.permissions.includes('admin'))) {
      return res.status(403).json({ error: `Forbidden: Requires ${permission} permission` });
    }
    next();
  };

  const isAdmin = (req: any, res: any, next: any) => {
    if (!req.user.permissions || !req.user.permissions.includes('admin')) {
      return res.status(403).json({ error: "Forbidden: Admin access required" });
    }
    next();
  };

  /** Entities from JWT (array) or legacy string (CSV / JSON array). */
  const parseUserEntities = (user: any): string[] => {
    const e = user?.entities;
    if (Array.isArray(e)) return e.map((x) => String(x).trim()).filter(Boolean);
    if (typeof e === "string") return parseCommaSeparatedList(e);
    return [];
  };

  const isAdminUser = (req: any) => !!req.user?.permissions?.includes("admin");

  const isDirectorUser = (req: any) =>
    !!req.user?.roles?.some((r: string) => r.toLowerCase() === "director") &&
    userHasDepartment(req.user.department, "management");

  const isPurchasingUser = (req: any) =>
    !!req.user?.roles?.some((r: string) => r.toLowerCase() === "purchasing");

  const isSomUser = (req: any) =>
    !!req.user?.roles?.some((r: string) => r.toLowerCase() === "som") &&
    userHasDepartment(req.user.department, "management");

  /**
   * Requires `X-Entity` header to match one of the user's allowed entities.
   * Sets `req.entityContext` for downstream handlers.
   */
  const requireEntityContext = (req: any, res: any, next: any) => {
    const raw = (req.headers["x-entity"] ?? req.headers["X-Entity"] ?? "").toString().trim();
    if (!raw) {
      return res.status(400).json({
        error: "Entity context is required. Send X-Entity header with the active entity code.",
      });
    }
    const allowed = parseUserEntities(req.user);
    const allowedByUpper = new Map(allowed.map((e) => [e.toUpperCase(), e]));
    const canonical = allowedByUpper.get(raw.toUpperCase()) || raw;
    if (!allowedByUpper.has(raw.toUpperCase())) {
      // Director/SOM (Management department) can access all entities.
      if (isDirectorUser(req) || isSomUser(req)) {
        req.entityContext = raw;
        next();
        return;
      }
      return res.status(403).json({ error: "You do not have access to this entity" });
    }
    req.entityContext = canonical;
    next();
  };

  /**
   * Row: workflow_requests row (may include template_category from join as template_category).
   * - Always: entity must match active entity context.
   * - Admin or director: any department within entity.
   * - Purchasing (on procurement workflows): any department within entity.
   * - Otherwise: request department must match one of user's departments.
   */
  const assertWorkflowRequestAccess = (req: any, res: any, requestRow: any, templateCategory?: string): boolean => {
    if (!requestRow) {
      res.status(404).json({ error: "Request not found" });
      return false;
    }
    if (!req.entityContext) {
      res.status(400).json({ error: "Entity context required" });
      return false;
    }
    const rowEntity = requestRow.entity ?? "";
    if (rowEntity !== req.entityContext) {
      res.status(403).json({ error: "Entity access denied" });
      return false;
    }
    const cat = templateCategory ?? requestRow.template_category ?? "";
    if (isAdminUser(req)) return true;
    if (isDirectorUser(req)) return true;
    if (cat === "procurement" && (isPurchasingUser(req) || isSomUser(req))) return true;
    if (userDepartmentsMatch(requestRow.department, req.user.department)) return true;
    res.status(403).json({ error: "Department access denied" });
    return false;
  };

  /** Public: LAN URLs for sharing; useful when the terminal is hidden. Set DISABLE_SERVER_ACCESS_URLS=1 to hide. */
  app.get("/api/server-access-urls", (_req, res) => {
    if (process.env.DISABLE_SERVER_ACCESS_URLS === "1") {
      return res.status(404).json({ error: "Not found" });
    }
    const lanHosts = getLanIPv4Addresses();
    const lanUrls = lanHosts.map((host) => `http://${host}:${PORT}`);
    res.json({
      port: PORT,
      localhostUrl: `http://localhost:${PORT}`,
      lanUrls,
      hostname: os.hostname(),
    });
  });

  // Auth Routes
  app.post("/api/login", async (req, res) => {
    const { username, password } = req.body;
    const sqlUser: any = await fetchSqlUsersettingByUsername(sqlPool, String(username ?? ""));
    if (!sqlUser || !sqlPasswordMatches(sqlUser.password, String(password ?? ""))) {
      return res.status(401).json({ error: "Invalid credentials" });
    }
    const roles = parseCommaSeparatedList(sqlUser.role);
    const entities = parseCommaSeparatedList(sqlUser.entities);
    const approval_limit_myr = await getUserApprovalLimitMyr(sqlPool, Number(sqlUser.id));
    const payload = {
      id: sqlUser.id,
      username: sqlUser.username,
      roles,
      department: sqlUser.department,
      entities,
      approval_limit_myr,
      permissions: getUserPermissions(roles),
    };
    const token = jwt.sign(payload, JWT_SECRET);
    approvalEventLog(`LOGIN success username=${approvalLogSanitize(String(username ?? ""), 120)} id=${sqlUser.id}`);
    return res.json({ token, user: payload });
  });

  app.post("/api/register", async (req, res) => {
    const { username, password, department, entities } = req.body;
    if (!username || !password) return res.status(400).json({ error: "Username and password are required" });
    const hashedPassword = bcrypt.hashSync(password, 10);
    const roleCsv = toCommaSeparated(["user"]);
    const entCsv = listToStoredCsv(entities || []);
    try {
      await sqlPool
        .request()
        .input("username", sql.NVarChar, username)
        .input("password", sql.NVarChar, hashedPassword)
        .input("role", sql.NVarChar(sql.MAX), roleCsv)
        .input("department", sql.NVarChar, department || "General")
        .input("entities", sql.NVarChar(sql.MAX), entCsv)
        .query(
          `INSERT INTO ${SQLSERVER_USERS_TABLE} (username, password, role, department, entities) VALUES (@username, @password, @role, @department, @entities)`
        );
      approvalEventLog(`REGISTER new_user username=${approvalLogSanitize(String(username), 120)}`);
      return res.json({ success: true });
    } catch {
      return res.status(400).json({ error: "Username already exists or registration failed" });
    }
  });

  app.get("/api/me", authenticate, (req: any, res) => {
    res.json(req.user);
  });

  // Saved signature for the current user (stored on network path as PNG)
  app.get("/api/me/signature", authenticate, async (req: any, res) => {
    try {
      const filePath = signatureFilePathForUser(req.user);
      const exists = fs.existsSync(filePath);
      if (!exists) return res.json({ exists: false });
      const buf = await fs.promises.readFile(filePath);
      const dataUrl = `data:image/png;base64,${buf.toString("base64")}`;
      return res.json({ exists: true, dataUrl });
    } catch (e: any) {
      return res.status(500).json({ error: "Failed to load saved signature", details: String(e?.message || e) });
    }
  });

  app.put("/api/me/signature", authenticate, async (req: any, res) => {
    try {
      const dataUrl = String(req.body?.dataUrl || "").trim();
      const buf = parsePngDataUrl(dataUrl);
      // Safety cap (~2.5MB) to allow high-resolution signatures while preventing huge uploads.
      if (buf.length > 2_500_000) return res.status(400).json({ error: "Signature image is too large (max 2.5MB)" });
      await fs.promises.mkdir(SIGNATURE_STORAGE_PATH, { recursive: true });
      const filePath = signatureFilePathForUser(req.user);
      await fs.promises.writeFile(filePath, buf);
      return res.json({ success: true });
    } catch (e: any) {
      return res.status(400).json({ error: String(e?.message || e || "Failed to save signature") });
    }
  });

  app.delete("/api/me/signature", authenticate, async (req: any, res) => {
    try {
      const filePath = signatureFilePathForUser(req.user);
      if (fs.existsSync(filePath)) {
        await fs.promises.unlink(filePath);
      }
      return res.json({ success: true });
    } catch (e: any) {
      return res.status(500).json({ error: "Failed to delete saved signature", details: String(e?.message || e) });
    }
  });

  // User Management (list; primary source = SQL Server `usersetting` when connected)
  app.get("/api/users", authenticate, async (req: any, res) => {
    const canViewUsers = req.user.permissions && (req.user.permissions.includes('manage_users') || req.user.permissions.includes('create_templates') || req.user.permissions.includes('admin'));
    if (!canViewUsers) {
      return res.status(403).json({ error: "Forbidden: Requires manage_users or create_templates permission" });
    }
    try {
      let rs;
      try {
        rs = await sqlPool.request().query(
          `SELECT id, username, role, department, approval_limit, designation FROM ${SQLSERVER_USERS_TABLE} ORDER BY username ASC`
        );
      } catch {
        try {
          rs = await sqlPool.request().query(
            `SELECT id, username, role, department, approval_limit FROM ${SQLSERVER_USERS_TABLE} ORDER BY username ASC`
          );
        } catch {
          rs = await sqlPool.request().query(
            `SELECT id, username, role, department FROM ${SQLSERVER_USERS_TABLE} ORDER BY username ASC`
          );
        }
      }
      const rows = (rs.recordset || []).map((u: any) => ({
        id: u.id,
        username: u.username,
        department: u.department,
        designation: u.designation != null ? String(u.designation).trim() || null : null,
        roles: parseCommaSeparatedList(u.role),
        approval_limit_myr: parseNullableNumber(u.approval_limit ?? u.approval_limit_myr),
      }));
      return res.json(rows);
    } catch (e: any) {
      console.warn("SQL user list failed:", e?.message || e);
      return res.status(500).json({ error: "Failed to load users" });
    }
  });

  app.post("/api/users", authenticate, hasPermission('manage_users'), async (req, res) => {
    const { username, password, roles, department } = req.body;
    if (!username || !password) return res.status(400).json({ error: "Username and password are required" });
    const hashedPassword = bcrypt.hashSync(password, 10);
    const roleCsv = listToStoredCsv(roles || ["user"]);
    try {
      await sqlPool
        .request()
        .input("username", sql.NVarChar, username)
        .input("password", sql.NVarChar, hashedPassword)
        .input("role", sql.NVarChar(sql.MAX), roleCsv)
        .input("department", sql.NVarChar, department || "General")
        .query(
          `INSERT INTO ${SQLSERVER_USERS_TABLE} (username, password, role, department) VALUES (@username, @password, @role, @department)`
        );
      return res.json({ success: true });
    } catch {
      return res.status(400).json({ error: "Username already exists or insert failed" });
    }
  });

  app.patch("/api/users/:id", authenticate, hasPermission('manage_users'), async (req, res) => {
    const { username, password, roles, department, designation } = req.body;
    const userId = Number(req.params.id);
    try {
      if (username) {
        await sqlPool.request().input("username", sql.NVarChar, username).input("id", sql.Int, userId).query(`UPDATE ${SQLSERVER_USERS_TABLE} SET username = @username WHERE id = @id`);
      }
      if (password) {
        const hashedPassword = bcrypt.hashSync(password, 10);
        await sqlPool.request().input("password", sql.NVarChar, hashedPassword).input("id", sql.Int, userId).query(`UPDATE ${SQLSERVER_USERS_TABLE} SET password = @password WHERE id = @id`);
      }
      if (roles) {
        await sqlPool
          .request()
          .input("role", sql.NVarChar(sql.MAX), listToStoredCsv(roles))
          .input("id", sql.Int, userId)
          .query(`UPDATE ${SQLSERVER_USERS_TABLE} SET role = @role WHERE id = @id`);
      }
      if (department) {
        await sqlPool.request().input("department", sql.NVarChar, department).input("id", sql.Int, userId).query(`UPDATE ${SQLSERVER_USERS_TABLE} SET department = @department WHERE id = @id`);
      }
      if (designation !== undefined) {
        const des = designation != null ? String(designation).trim() : "";
        await sqlPool
          .request()
          .input("designation", sql.NVarChar, des || null)
          .input("id", sql.Int, userId)
          .query(`UPDATE ${SQLSERVER_USERS_TABLE} SET designation = @designation WHERE id = @id`);
      }
      return res.json({ success: true });
    } catch (err: any) {
      return res.status(400).json({ error: err.message });
    }
  });

  app.delete("/api/users/:id", authenticate, hasPermission('manage_users'), async (req: any, res) => {
    return res.status(403).json({ error: "Delete operations are disabled in this system." });
  });

  app.patch("/api/users/:id/role", authenticate, hasPermission('manage_users'), async (req, res) => {
    const { roles, department } = req.body;
    const id = Number(req.params.id);
    if (roles && department) {
      await sqlPool
        .request()
        .input("role", sql.NVarChar(sql.MAX), listToStoredCsv(roles))
        .input("department", sql.NVarChar, department)
        .input("id", sql.Int, id)
        .query(`UPDATE ${SQLSERVER_USERS_TABLE} SET role = @role, department = @department WHERE id = @id`);
    } else if (roles) {
      await sqlPool
        .request()
        .input("role", sql.NVarChar(sql.MAX), listToStoredCsv(roles))
        .input("id", sql.Int, id)
        .query(`UPDATE ${SQLSERVER_USERS_TABLE} SET role = @role WHERE id = @id`);
    } else if (department) {
      await sqlPool.request().input("department", sql.NVarChar, department).input("id", sql.Int, id).query(`UPDATE ${SQLSERVER_USERS_TABLE} SET department = @department WHERE id = @id`);
    }
    return res.json({ success: true });
  });

  /** Users who may be picked as approver for the active entity context. */
  app.get("/api/users/eligible-approvers", authenticate, requireEntityContext, async (req: any, res) => {
    try {
      const activeEnt = String(req.entityContext || "").trim().toUpperCase();
      const requesterRow = await fetchSqlUsersettingById(sqlPool, Number(req.user?.id));
      const requesterDepartment =
        requesterRow && requesterRow.department != null
          ? String(requesterRow.department)
          : String(req.user?.department || "");
      let rs;
      try {
        rs = await sqlPool.request().query(
          `SELECT id, username, role, department, entities FROM ${SQLSERVER_USERS_TABLE} ORDER BY username ASC`
        );
      } catch {
        return res.status(500).json({ error: "Failed to load users" });
      }
      const out = (rs.recordset || [])
        .filter(
          (u: any) =>
            Number(u.id) !== Number(req.user.id) &&
            userRowHasEntityAccess(u, activeEnt) &&
            userRowHasApproverRole(u) &&
            userDepartmentsMatch(u.department, requesterDepartment)
        )
        .map((u: any) => ({ id: u.id, username: u.username, department: u.department || "" }));
      return res.json(out);
    } catch (e: any) {
      return res.status(500).json({ error: e?.message || "Failed to load approvers" });
    }
  });

  /** Rows from `cost_center` on the SQL catalog (source of truth for dropdowns). */
  app.get("/api/cost-centers", authenticate, requireEntityContext, async (req: any, res) => {
    const canManage =
      req.user?.permissions?.includes("manage_users") || req.user?.permissions?.includes("admin");
    const includeInactive = canManage && String(req.query.include_inactive || "") === "1";
    const ent = String(req.entityContext || "").trim().toUpperCase();
    if (!ent) return res.status(400).json({ error: "entity is required" });
    const exposeDetails = process.env.NODE_ENV !== "production";
    try {
      await ensureCostCentersTable(sqlPool);
    } catch (e: any) {
      console.error("GET /api/cost-centers schema ensure:", formatSqlError(e));
      return res.status(500).json({
        error:
          "Cost centers catalog is unavailable (database could not create or update dbo.cost_center). Check SQL permissions and server logs.",
        ...(exposeDetails ? { details: formatSqlError(e) } : {}),
      });
    }
    try {
      const tableInfo = await resolveCostCenterTableInfo(sqlPool);
      const idProjection = tableInfo.hasId ? "id" : "NULL AS id";
      const q = includeInactive
        ? `SELECT ${idProjection}, entity, code, name, gl_account, approval, status, created_at FROM ${tableInfo.qualifiedName}
           WHERE UPPER(LTRIM(RTRIM(entity))) = @entity
           ORDER BY code ASC`
        : `SELECT ${idProjection}, entity, code, name, gl_account, approval, status, created_at FROM ${tableInfo.qualifiedName}
           WHERE UPPER(LTRIM(RTRIM(entity))) = @entity AND status = 1
           ORDER BY code ASC`;
      const rs = await sqlPool.request().input("entity", sql.NVarChar, ent).query(q);
      const rows = (rs.recordset || []).map((r: any) => {
        const codeRaw = r.code ?? r.Code;
        const nameRaw = r.name ?? r.Name;
        return {
          id: r.id,
          entity: r.entity,
          code: codeRaw != null ? String(codeRaw).trim() : "",
          name: nameRaw != null ? String(nameRaw).trim() : "",
          gl_account: r.gl_account ?? r.GL_Account ?? null,
          approval: r.approval ?? r.Approval ?? null,
          status: !!r.status,
          created_at: r.created_at,
        };
      });
      return res.json(rows);
    } catch (e: any) {
      console.error("GET /api/cost-centers:", e?.message || e);
      return res.status(500).json({
        error: "Failed to load cost centers",
        ...(exposeDetails ? { details: String(e?.message || e) } : {}),
      });
    }
  });

  app.post("/api/cost-centers", authenticate, requireEntityContext, hasPermission("manage_users"), async (req: any, res) => {
    const entity = String(req.entityContext || "").trim().toUpperCase();
    const bodyEntity = String(req.body?.entity || "").trim().toUpperCase();
    if (bodyEntity && bodyEntity !== entity) {
      return res.status(400).json({ error: "entity in body must match active entity (X-Entity header)" });
    }
    const code = String(req.body?.code ?? "").trim();
    const name = String(req.body?.name ?? "").trim();
    const gl_account = req.body?.gl_account != null ? String(req.body.gl_account).trim() : "";
    const approval = req.body?.approval != null ? String(req.body.approval).trim() : "";
    let statusBit = 1;
    if (req.body?.status !== undefined && req.body?.status !== null) {
      const s = Number(req.body.status);
      if (s !== 0 && s !== 1) return res.status(400).json({ error: "status must be 0 or 1" });
      statusBit = s;
    }
    if (!entity || !code || !name) return res.status(400).json({ error: "entity, code and name are required" });
    try {
      const tableInfo = await resolveCostCenterTableInfo(sqlPool);
      const outputClause = tableInfo.hasId ? "OUTPUT INSERTED.id AS id" : "";
      const ins = await sqlPool
        .request()
        .input("entity", sql.NVarChar, entity)
        .input("code", sql.NVarChar, code)
        .input("name", sql.NVarChar, name)
        .input("gl_account", sql.NVarChar, gl_account || null)
        .input("approval", sql.NVarChar, approval || null)
        .input("status", sql.Bit, statusBit)
        .query(`
          INSERT INTO ${tableInfo.qualifiedName} (entity, code, name, gl_account, approval, status)
          ${outputClause}
          VALUES (@entity, @code, @name, @gl_account, @approval, @status)
        `);
      const id = ins.recordset?.[0]?.id;
      return res.json({ id });
    } catch (e: any) {
      const msg = String(e?.message || e || "");
      if (/UQ_cost_centers|duplicate|unique/i.test(msg)) {
        return res.status(400).json({ error: "A cost center with this code already exists for this entity" });
      }
      return res.status(400).json({ error: msg || "Failed to create cost center" });
    }
  });

  app.patch("/api/cost-centers/:id", authenticate, requireEntityContext, hasPermission("manage_users"), async (req: any, res) => {
    const id = Number(req.params.id);
    const ent = String(req.entityContext || "").trim().toUpperCase();
    const bodyEntity = String(req.body?.entity || "").trim().toUpperCase();
    if (!Number.isFinite(id) || id <= 0) return res.status(400).json({ error: "Invalid id" });
    if (!ent) return res.status(400).json({ error: "entity is required" });
    if (bodyEntity && bodyEntity !== ent) {
      return res.status(400).json({ error: "entity in body must match active entity (X-Entity header)" });
    }
    const { code, name, gl_account, approval, status } = req.body || {};
    try {
      const tableInfo = await resolveCostCenterTableInfo(sqlPool);
      if (!tableInfo.hasId) return res.status(400).json({ error: "Table cost_center does not expose an id column." });
      const cur = await sqlPool
        .request()
        .input("id", sql.Int, id)
        .input("entity", sql.NVarChar, ent)
        .query(`SELECT TOP 1 * FROM ${tableInfo.qualifiedName} WHERE id = @id AND UPPER(LTRIM(RTRIM(entity))) = @entity`);
      const row = cur.recordset?.[0];
      if (!row) return res.status(404).json({ error: "Cost center not found" });
      const nextCode = code !== undefined ? String(code).trim() : String(row.code);
      const nextName = name !== undefined ? String(name).trim() : String(row.name);
      const nextGl = gl_account !== undefined ? (String(gl_account).trim() || null) : row.gl_account;
      const nextAppr = approval !== undefined ? (String(approval).trim() || null) : row.approval;
      let nextStatus = !!row.status;
      if (status !== undefined && status !== null) {
        const s = Number(status);
        if (s !== 0 && s !== 1) return res.status(400).json({ error: "status must be 0 or 1" });
        nextStatus = s === 1;
      }
      if (!nextCode || !nextName) return res.status(400).json({ error: "code and name cannot be empty" });
      await sqlPool
        .request()
        .input("id", sql.Int, id)
        .input("entity", sql.NVarChar, ent)
        .input("code", sql.NVarChar, nextCode)
        .input("name", sql.NVarChar, nextName)
        .input("gl_account", sql.NVarChar, nextGl)
        .input("approval", sql.NVarChar, nextAppr)
        .input("status", sql.Bit, nextStatus ? 1 : 0)
        .query(
          `UPDATE ${tableInfo.qualifiedName}
           SET code = @code, name = @name, gl_account = @gl_account, approval = @approval, status = @status
           WHERE id = @id AND UPPER(LTRIM(RTRIM(entity))) = @entity`
        );
      return res.json({ success: true });
    } catch (e: any) {
      const msg = String(e?.message || e || "");
      if (/UQ_cost_centers|duplicate|unique/i.test(msg)) {
        return res.status(400).json({ error: "A cost center with this code already exists for this entity" });
      }
      return res.status(400).json({ error: msg || "Update failed" });
    }
  });

  app.delete("/api/cost-centers/:id", authenticate, hasPermission("manage_users"), async (req: any, res) => {
    return res.status(403).json({ error: "Delete operations are disabled in this system." });
  });

  /** SR spare-for locations master list (entity scoped; active rows only unless admin passes include_inactive=1). */
  app.get("/api/spare-locations", authenticate, async (req: any, res) => {
    const canManage =
      req.user?.permissions?.includes("manage_users") || req.user?.permissions?.includes("admin");
    const includeInactive = canManage && String(req.query.include_inactive || "") === "1";
    const ent = String(req.query.entity || req.entityContext || "").trim().toUpperCase();
    if (!ent) return res.status(400).json({ error: "entity is required" });
    try {
      const q = includeInactive
        ? `SELECT id, entity, code, name, status, created_at
           FROM dbo.spare_locations
           WHERE UPPER(LTRIM(RTRIM(entity))) = @entity
           ORDER BY code ASC`
        : `SELECT id, entity, code, name, status, created_at
           FROM dbo.spare_locations
           WHERE UPPER(LTRIM(RTRIM(entity))) = @entity AND status = 1
           ORDER BY code ASC`;
      const rs = await sqlPool.request().input("entity", sql.NVarChar, ent).query(q);
      const rows = (rs.recordset || []).map((r: any) => ({
        id: r.id,
        entity: r.entity,
        code: r.code,
        name: r.name,
        status: !!r.status,
        created_at: r.created_at,
      }));
      return res.json(rows);
    } catch (e: any) {
      console.error("GET /api/spare-locations:", e?.message || e);
      return res.status(500).json({ error: "Failed to load spare locations" });
    }
  });

  app.post("/api/spare-locations", authenticate, hasPermission("manage_users"), async (req: any, res) => {
    const entity = String(req.body?.entity ?? "").trim().toUpperCase();
    const code = String(req.body?.code ?? "").trim();
    const name = String(req.body?.name ?? "").trim();
    let statusBit = 1;
    if (req.body?.status !== undefined && req.body?.status !== null) {
      const s = Number(req.body.status);
      if (s !== 0 && s !== 1) return res.status(400).json({ error: "status must be 0 or 1" });
      statusBit = s;
    }
    if (!entity || !code || !name) return res.status(400).json({ error: "entity, code and name are required" });
    try {
      const ins = await sqlPool
        .request()
        .input("entity", sql.NVarChar, entity)
        .input("code", sql.NVarChar, code)
        .input("name", sql.NVarChar, name)
        .input("status", sql.Bit, statusBit)
        .query(`
          INSERT INTO dbo.spare_locations (entity, code, name, status)
          OUTPUT INSERTED.id AS id
          VALUES (@entity, @code, @name, @status)
        `);
      const id = ins.recordset?.[0]?.id;
      return res.json({ id });
    } catch (e: any) {
      const msg = String(e?.message || e || "");
      if (/UQ_spare_locations_entity_code|duplicate|unique/i.test(msg)) {
        return res.status(400).json({ error: "A spare location with this code already exists for this entity" });
      }
      return res.status(400).json({ error: msg || "Failed to create spare location" });
    }
  });

  app.patch("/api/spare-locations/:id", authenticate, hasPermission("manage_users"), async (req: any, res) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id) || id <= 0) return res.status(400).json({ error: "Invalid id" });
    const { entity, code, name, status } = req.body || {};
    try {
      const cur = await sqlPool.request().input("id", sql.Int, id).query(`SELECT TOP 1 * FROM dbo.spare_locations WHERE id = @id`);
      const row = cur.recordset?.[0];
      if (!row) return res.status(404).json({ error: "Spare location not found" });
      const nextEntity = entity !== undefined ? String(entity).trim().toUpperCase() : String(row.entity).trim().toUpperCase();
      const nextCode = code !== undefined ? String(code).trim() : String(row.code);
      const nextName = name !== undefined ? String(name).trim() : String(row.name);
      let nextStatus = !!row.status;
      if (status !== undefined && status !== null) {
        const s = Number(status);
        if (s !== 0 && s !== 1) return res.status(400).json({ error: "status must be 0 or 1" });
        nextStatus = s === 1;
      }
      if (!nextEntity || !nextCode || !nextName) {
        return res.status(400).json({ error: "entity, code and name cannot be empty" });
      }
      await sqlPool
        .request()
        .input("id", sql.Int, id)
        .input("entity", sql.NVarChar, nextEntity)
        .input("code", sql.NVarChar, nextCode)
        .input("name", sql.NVarChar, nextName)
        .input("status", sql.Bit, nextStatus ? 1 : 0)
        .query(
          `UPDATE dbo.spare_locations
           SET entity = @entity, code = @code, name = @name, status = @status
           WHERE id = @id`
        );
      return res.json({ success: true });
    } catch (e: any) {
      const msg = String(e?.message || e || "");
      if (/UQ_spare_locations_entity_code|duplicate|unique/i.test(msg)) {
        return res.status(400).json({ error: "A spare location with this code already exists for this entity" });
      }
      return res.status(400).json({ error: msg || "Update failed" });
    }
  });

  app.delete("/api/spare-locations/:id", authenticate, hasPermission("manage_users"), async (req: any, res) => {
    return res.status(403).json({ error: "Delete operations are disabled in this system." });
  });

  /**
   * Upsert `entity_approver_registry` (per-entity cap + optional GCCM picker flag).
   * Omit `approval_limit_myr` or `selectable_by_requestor` from the body to leave that column unchanged on update.
   */
  app.put("/api/users/:id/entity-approval-limit", authenticate, hasPermission("manage_users"), async (req: any, res) => {
    const userId = Number(req.params.id);
    const { entity, approval_limit_myr, selectable_by_requestor } = req.body;
    const ent = String(entity ?? "")
      .trim()
      .toUpperCase();
    if (!ent) return res.status(400).json({ error: "entity is required (e.g. GCCM)" });
    const limSet = Object.prototype.hasOwnProperty.call(req.body, "approval_limit_myr");
    const pickSet = Object.prototype.hasOwnProperty.call(req.body, "selectable_by_requestor");
    const lim = approval_limit_myr === undefined || approval_limit_myr === "" ? null : Number(approval_limit_myr);
    if (limSet && lim !== null && (!Number.isFinite(lim) || lim < 0)) {
      return res.status(400).json({ error: "approval_limit_myr must be a non-negative number or null to clear" });
    }
    const pickBit = pickSet && !!selectable_by_requestor ? 1 : 0;
    try {
      await sqlPool
        .request()
        .input("user_id", sql.Int, userId)
        .input("entity", sql.NVarChar, ent)
        .input("limSet", sql.Bit, limSet ? 1 : 0)
        .input("lim", sql.Decimal(18, 2), limSet ? lim : null)
        .input("pickSet", sql.Bit, pickSet ? 1 : 0)
        .input("pick", sql.Bit, pickBit)
        .query(`
          MERGE dbo.entity_approver_registry AS t
          USING (SELECT @entity AS entity, @user_id AS user_id) AS s
          ON UPPER(LTRIM(RTRIM(t.entity))) = s.entity AND t.user_id = s.user_id
          WHEN MATCHED THEN UPDATE SET
            approval_limit_myr = CASE WHEN @limSet = 1 THEN @lim ELSE t.approval_limit_myr END,
            selectable_by_requestor = CASE WHEN @pickSet = 1 THEN @pick ELSE t.selectable_by_requestor END,
            active = 1
          WHEN NOT MATCHED THEN INSERT (entity, user_id, selectable_by_requestor, approval_limit_myr, active)
          VALUES (
            @entity,
            @user_id,
            CASE WHEN @pickSet = 1 THEN @pick ELSE 0 END,
            CASE WHEN @limSet = 1 THEN @lim ELSE NULL END,
            1
          );
        `);
      if (limSet) {
        await sqlPool
          .request()
          .input("user_id", sql.Int, userId)
          .input("entity", sql.NVarChar, ent)
          .input("lim", sql.Decimal(18, 2), lim)
          .query(`
            MERGE dbo.user_entity_approval_limits AS t
            USING (SELECT @user_id AS user_id, @entity AS entity) AS s
            ON t.user_id = s.user_id AND UPPER(LTRIM(RTRIM(t.entity))) = s.entity
            WHEN MATCHED THEN UPDATE SET approval_limit_myr = @lim
            WHEN NOT MATCHED THEN INSERT (user_id, entity, approval_limit_myr) VALUES (@user_id, @entity, @lim);
          `);
      }
      approvalEventLog(
        `ENTITY_APPROVER_REGISTRY ${formatActor(req)} target_user_id=${userId} entity=${ent} lim_set=${limSet ? 1 : 0} limit_myr=${limSet ? (lim === null ? "null" : String(lim)) : "unchanged"} pick_set=${pickSet ? 1 : 0} selectable=${pickSet ? String(pickBit) : "unchanged"}`
      );
      return res.json({ success: true });
    } catch (e: any) {
      return res.status(400).json({ error: e?.message || "Update failed" });
    }
  });

  // Role Management (Admin Only)
  app.get("/api/roles", authenticate, async (req, res) => {
    try {
      const rs = await sqlPool.request().query("SELECT id, name, permissions, created_at FROM custom_roles ORDER BY name ASC");
      const roles = (rs.recordset || []).map((r: any) => ({
        ...r,
        permissions: parseCommaSeparatedList(r.permissions),
      }));
      return res.json(roles);
    } catch (e: any) {
      console.warn("SQL roles fetch failed:", e?.message || e);
      return res.status(500).json({ error: "Failed to load roles" });
    }
  });

  app.post("/api/roles", authenticate, isAdmin, async (req, res) => {
    const { name, permissions } = req.body;
    if (!name) return res.status(400).json({ error: "Role name is required" });
    const roleName = String(name).toLowerCase();
    const permsCsv = listToStoredCsv(permissions || []);
    try {
      await sqlPool
        .request()
        .input("name", sql.NVarChar, roleName)
        .input("permissions", sql.NVarChar(sql.MAX), permsCsv)
        .query("INSERT INTO custom_roles (name, permissions) VALUES (@name, @permissions)");
      const created = await sqlPool.request().input("name", sql.NVarChar, roleName).query("SELECT TOP 1 id, name, permissions FROM custom_roles WHERE name = @name");
      const row: any = created.recordset?.[0];
      await refreshRolePermissionsCache(sqlPool);
      return res.json({
        id: row?.id,
        name: row?.name || roleName,
        permissions: parseCommaSeparatedList(row?.permissions),
      });
    } catch (err) {
      return res.status(400).json({ error: "Role already exists" });
    }
  });

  app.patch("/api/roles/:id", authenticate, isAdmin, async (req, res) => {
    const { name, permissions } = req.body;
    const roleId = req.params.id;
    try {
      if (name) {
        await sqlPool.request().input("name", sql.NVarChar, String(name).toLowerCase()).input("id", sql.Int, Number(roleId)).query("UPDATE custom_roles SET name = @name WHERE id = @id");
      }
      if (permissions !== undefined) {
        await sqlPool
          .request()
          .input("permissions", sql.NVarChar(sql.MAX), listToStoredCsv(permissions))
          .input("id", sql.Int, Number(roleId))
          .query("UPDATE custom_roles SET permissions = @permissions WHERE id = @id");
      }
      await refreshRolePermissionsCache(sqlPool);
      return res.json({ success: true });
    } catch (err) {
      return res.status(400).json({ error: "Error updating role" });
    }
  });

  app.delete("/api/roles/:id", authenticate, isAdmin, async (req, res) => {
    return res.status(403).json({ error: "Delete operations are disabled in this system." });
  });

  // Workflow / request APIs (SQL Server; join users via SQLSERVER_USERS_TABLE)
  const userJoinSql = SQLSERVER_USERS_TABLE;

  app.get("/api/workflows", authenticate, async (req: any, res) => {
    const canSeeAll = req.user.permissions && (
      req.user.permissions.includes("admin") ||
      req.user.permissions.includes("create_templates") ||
      req.user.permissions.includes("approve_templates")
    );
    const userIsDirector = isDirectorUser(req);

    try {
      let rs;
      if (canSeeAll || userIsDirector) {
        rs = await sqlPool.request().query(`
            SELECT w.*, u.username AS creator_name
            FROM workflows w
            INNER JOIN ${userJoinSql} u ON w.creator_id = u.id
          `);
      } else {
        rs = await sqlPool.request().input("creatorId", sql.Int, req.user.id).query(`
            SELECT w.*, u.username AS creator_name
            FROM workflows w
            INNER JOIN ${userJoinSql} u ON w.creator_id = u.id
            WHERE w.creator_id = @creatorId OR (w.status = N'approved' AND ISNULL(w.is_active, 1) = 1)
          `);
      }
      const workflows = rs.recordset || [];
      return res.json(
        workflows.map((w: any) => ({
          ...w,
          steps: JSON.parse(w.steps || "[]"),
          table_columns: w.table_columns ? JSON.parse(w.table_columns) : [],
          attachments_required: !!w.attachments_required,
        }))
      );
    } catch (e: any) {
      console.error("SQL GET /api/workflows:", e?.message || e);
      return res.status(500).json({ error: "Failed to load workflows from database" });
    }
  });

  app.post("/api/workflows", authenticate, hasPermission("create_templates"), async (req: any, res) => {
    let { name, category, steps, table_columns, attachments_required } = req.body;
    if (category === "procurement" && isPOName(name)) {
      steps = FIXED_PO_STEPS_FULL;
    } else if (category === "procurement" && isPRName(name)) {
      steps = FIXED_PR_STEPS;
    } else if (category === "procurement" && isSRName(name)) {
      steps = FIXED_SR_STEPS;
    }

    try {
      const ins = await sqlPool
        .request()
        .input("creator_id", sql.Int, req.user.id)
        .input("name", sql.NVarChar, name)
        .input("category", sql.NVarChar, category || "general")
        .input("steps", sql.NVarChar(sql.MAX), JSON.stringify(steps))
        .input("table_columns", sql.NVarChar(sql.MAX), JSON.stringify(table_columns || []))
        .input("attachments_required", sql.Bit, attachments_required ? 1 : 0)
        .query(`
            INSERT INTO workflows (creator_id, name, category, steps, table_columns, attachments_required)
            OUTPUT INSERTED.id AS id
            VALUES (@creator_id, @name, @category, @steps, @table_columns, @attachments_required)
          `);
      const id = ins.recordset?.[0]?.id;
      approvalEventLog(
        `WORKFLOW_CREATE ${formatActor(req)} workflow_id=${id} name=${approvalLogSanitize(String(name || ""), 200)} category=${approvalLogSanitize(String(category || ""), 80)}`
      );
      return res.json({ id });
    } catch (e: any) {
      console.error("SQL POST /api/workflows:", e?.message || e);
      return res.status(400).json({ error: e?.message || "Failed to create workflow" });
    }
  });

  /** Template-level `attachments` table was removed; files for requests live on disk (see request_attachments). */
  app.get("/api/workflows/:id/attachments", authenticate, async (_req, res) => {
    return res.json([]);
  });

  app.patch("/api/workflows/:id/status", authenticate, async (req: any, res) => {
    if (!req.user.permissions || (!req.user.permissions.includes("approve_templates") && !req.user.permissions.includes("admin"))) {
      return res.status(403).json({ error: "You do not have permission to approve templates" });
    }
    const { status } = req.body;
    try {
      await sqlPool
        .request()
        .input("status", sql.NVarChar, status)
        .input("id", sql.Int, Number(req.params.id))
        .query("UPDATE workflows SET status = @status WHERE id = @id");
      approvalEventLog(
        `WORKFLOW_STATUS ${formatActor(req)} workflow_id=${Number(req.params.id)} status=${approvalLogSanitize(String(status ?? ""), 40)}`
      );
      return res.json({ success: true });
    } catch (e: any) {
      return res.status(400).json({ error: e?.message || "Update failed" });
    }
  });

  app.patch("/api/workflows/:id", authenticate, async (req: any, res) => {
    let { name, category, steps, table_columns, attachments_required } = req.body;
    const workflowId = req.params.id;

    try {
      const wfRs = await sqlPool.request().input("id", sql.Int, Number(workflowId)).query("SELECT TOP 1 * FROM workflows WHERE id = @id");
      const workflow: any = wfRs.recordset?.[0];
      if (!workflow) return res.status(404).json({ error: "Workflow not found" });
      const finalName = name || workflow.name;
      const finalCategory = category || workflow.category;
      if (finalCategory === "procurement" && isPOName(finalName)) {
        steps = FIXED_PO_STEPS_FULL;
      } else if (finalCategory === "procurement" && isPRName(finalName)) {
        steps = FIXED_PR_STEPS;
      } else if (finalCategory === "procurement" && isSRName(finalName)) {
        steps = FIXED_SR_STEPS;
      }
      const canEdit =
        req.user.permissions &&
        (req.user.permissions.includes("admin") ||
          req.user.permissions.includes("create_templates") ||
          (req.user.permissions.includes("approve_templates") && workflow.status === "pending") ||
          req.user.id === workflow.creator_id);
      if (!canEdit) {
        return res.status(403).json({ error: "You do not have permission to edit this template" });
      }
      const wid = Number(workflowId);
      if (name !== undefined) {
        await sqlPool.request().input("name", sql.NVarChar, name).input("id", sql.Int, wid).query("UPDATE workflows SET name = @name WHERE id = @id");
      }
      if (category !== undefined) {
        await sqlPool.request().input("category", sql.NVarChar, category).input("id", sql.Int, wid).query("UPDATE workflows SET category = @category WHERE id = @id");
      }
      if (steps !== undefined) {
        await sqlPool
          .request()
          .input("steps", sql.NVarChar(sql.MAX), JSON.stringify(steps))
          .input("id", sql.Int, wid)
          .query("UPDATE workflows SET steps = @steps WHERE id = @id");
      }
      if (table_columns !== undefined) {
        await sqlPool
          .request()
          .input("table_columns", sql.NVarChar(sql.MAX), JSON.stringify(table_columns))
          .input("id", sql.Int, wid)
          .query("UPDATE workflows SET table_columns = @table_columns WHERE id = @id");
      }
      if (attachments_required !== undefined) {
        await sqlPool
          .request()
          .input("attachments_required", sql.Bit, attachments_required ? 1 : 0)
          .input("id", sql.Int, wid)
          .query("UPDATE workflows SET attachments_required = @attachments_required WHERE id = @id");
      }
      approvalEventLog(`WORKFLOW_UPDATE ${formatActor(req)} workflow_id=${Number(workflowId)}`);
      return res.json({ success: true });
    } catch (err: any) {
      return res.status(400).json({ error: err.message });
    }
  });

  app.delete("/api/workflows/:id", authenticate, async (req: any, res) => {
    return res.status(403).json({ error: "Delete operations are disabled in this system." });
  });

  app.get("/api/workflow-requests", authenticate, async (req: any, res) => {
    try {
      const reqSql = sqlPool.request();
      const requestedEntity = String(req.headers["x-entity"] ?? req.headers["X-Entity"] ?? "").trim();
      const allowedEntities = parseUserEntities(req.user).map((e) => e.trim()).filter(Boolean);
      const allowedSetUpper = new Set(allowedEntities.map((e) => e.toUpperCase()));
      const canViewAllEntities = isAdminUser(req) || isDirectorUser(req) || isSomUser(req);
      let entityClause = "";
      if (requestedEntity) {
        if (!canViewAllEntities && !allowedSetUpper.has(requestedEntity.toUpperCase())) {
          return res.status(403).json({ error: "You do not have access to this entity" });
        }
        entityClause = " AND r.entity = @entity";
        reqSql.input("entity", sql.NVarChar, requestedEntity);
      } else if (!canViewAllEntities) {
        if (allowedEntities.length === 0) return res.json([]);
        const params: string[] = [];
        allowedEntities.forEach((ent, idx) => {
          const key = `ent_${idx}`;
          reqSql.input(key, sql.NVarChar, ent);
          params.push(`@${key}`);
        });
        entityClause = ` AND r.entity IN (${params.join(", ")})`;
      }
      let deptClause = "";
      if (!isAdminUser(req) && !isDirectorUser(req) && !isSomUser(req)) {
        const userDepartments = parseCommaSeparatedList(req.user.department);
        if (userDepartments.length === 0) return res.json([]);
        const deptParams: string[] = [];
        userDepartments.forEach((d, idx) => {
          const key = `dept_${idx}`;
          reqSql.input(key, sql.NVarChar, d);
          deptParams.push(`@${key}`);
        });
        deptClause = ` AND r.department IN (${deptParams.join(", ")})`;
      }
      const rs = await reqSql.query(`
          SELECT r.*,
            po_link.formatted_id AS linked_po_formatted_id,
            po_link.status AS linked_po_status,
            COALESCE(NULLIF(LTRIM(RTRIM(CAST(r.requester_name AS NVARCHAR(255)))), N''), u.username) AS computed_requester_name,
            u.designation AS requester_designation,
            w.name AS template_name,
            COALESCE(r.request_steps, w.steps) AS template_steps, w.table_columns AS table_columns,
            w.attachments_required AS attachments_required, w.category AS category
          FROM workflow_requests r
          INNER JOIN ${userJoinSql} u ON r.requester_id = u.id
          INNER JOIN workflows w ON r.template_id = w.id
          LEFT JOIN workflow_requests po_link ON po_link.id = r.converted_po_request_id
          WHERE 1=1${entityClause}${deptClause}
          ORDER BY r.created_at DESC
        `);
      const requests = rs.recordset || [];
      return res.json(
        requests.map((r: any) => {
          const { computed_requester_name, ...row } = r;
          return {
            ...row,
            requester_name: computed_requester_name ?? row.requester_name,
            template_steps: JSON.parse(row.template_steps),
            table_columns: row.table_columns ? JSON.parse(row.table_columns) : [],
            line_items: row.line_items ? JSON.parse(row.line_items) : [],
            attachments_required: !!row.attachments_required,
          };
        })
      );
    } catch (e: any) {
      console.error("SQL GET /api/workflow-requests:", e?.message || e);
      return res.status(500).json({ error: "Failed to load workflow requests" });
    }
  });

  app.post("/api/workflow-requests", authenticate, requireEntityContext, async (req: any, res) => {
    const {
      template_id,
      title,
      details,
      line_items,
      attachments,
      tax_rate,
      discount_rate,
      currency,
      cost_center,
      section,
      entity,
      requester_signature,
      assigned_approver_id,
    } = req.body;
    const requesterSigned = bodyIndicatesRequesterSigned(req.body);
    const requesterSigStored =
      typeof requester_signature === "string" && requester_signature.trim().length > 0 ? requester_signature.trim() : null;

    const bodyEntity = (entity ?? "").toString().trim() || null;
    if (bodyEntity !== req.entityContext) {
      return res.status(400).json({ error: "Request entity must match the active entity (X-Entity header)" });
    }

    let template: any;
    try {
      const tr = await sqlPool
        .request()
        .input("id", sql.Int, Number(template_id))
        .query("SELECT TOP 1 status, name, steps, category FROM workflows WHERE id = @id");
      template = tr.recordset?.[0];
    } catch (e: any) {
      return res.status(500).json({ error: e?.message || "Failed to load template" });
    }

    if (!template || template.status !== "approved") {
      return res.status(400).json({ error: "Cannot use an unapproved workflow template" });
    }

    const currencyTrimmed = currency != null ? String(currency).trim() : "";
    const isPR = isPRName(template.name);
    const isPO = isPOName(template.name);
    const isSR = isSRName(template.name);
    if (template.category === "procurement" && (isPR || isPO || isSR) && !currencyTrimmed) {
      return res.status(400).json({
        error: "Currency is required for purchase requests, stock requisitions, and purchase orders",
      });
    }
    const sectionTrimmedRaw = String(section ?? "").trim();
    const sectionTrimmed = sectionTrimmedRaw.toUpperCase() === "NA" ? "" : sectionTrimmedRaw;

    let formatted_id: string | null = null;
    const entityCode = req.entityContext;
    if ((isPR || isPO || isSR) && entityCode) {
      const year = new Date().getFullYear().toString().slice(-2);
      const prefixMap: { [key: string]: string } = {
        CCI: "CCI",
        GCBCM: "CM",
        GCCM: "GC",
        GCBCS: "SG",
        ACI: "ACI",
      };
      const entityPrefix = prefixMap[entityCode] || entityCode.toUpperCase();
      const typePrefix = isPR ? "PR" : isPO ? "PO" : "SR";
      const pattern = `${typePrefix}${entityPrefix}${year}-%`;

      const lr = await sqlPool
        .request()
        .input("entity", sql.NVarChar, entityCode)
        .input("pattern", sql.NVarChar, pattern)
        .query(
          "SELECT TOP 1 formatted_id FROM workflow_requests WHERE entity = @entity AND formatted_id LIKE @pattern ORDER BY formatted_id DESC"
        );
      const lastRequest = lr.recordset?.[0];

      let nextNum = 1;
      if (lastRequest && lastRequest.formatted_id) {
        const parts = lastRequest.formatted_id.split("-");
        const lastNum = parseInt(parts[parts.length - 1]);
        if (!isNaN(lastNum)) nextNum = lastNum + 1;
      }
      formatted_id = `${typePrefix}${entityPrefix}${year}-${nextNum.toString().padStart(5, "0")}`;
    }

    let requestStepsJson: string | null = null;
    const discValCreate = discount_rate !== undefined ? Number(discount_rate) : 0;
    const suggestedSupplierRaw =
      req.body.suggested_supplier != null ? String(req.body.suggested_supplier).trim() : "";
    let suggestedSupplierSql: string | null = null;
    if (isPR) {
      if (!suggestedSupplierRaw) {
        return res.status(400).json({ error: "Suggested supplier is required for purchase requests" });
      }
      suggestedSupplierSql = suggestedSupplierRaw;
    }
    if (isPOName(template.name)) {
      requestStepsJson = JSON.stringify(
        buildPoRequestSteps(
          computeRequestTotalMyr(line_items || [], tax_rate !== undefined ? tax_rate : 0.18, currencyTrimmed, discValCreate)
        )
      );
    } else if (isPRName(template.name)) {
      requestStepsJson = JSON.stringify(FIXED_PR_STEPS);
    } else if (isSRName(template.name)) {
      requestStepsJson = JSON.stringify(FIXED_SR_STEPS);
    } else {
      try {
        requestStepsJson = JSON.stringify(JSON.parse(template.steps || "[]"));
      } catch {
        requestStepsJson = JSON.stringify([]);
      }
    }

    const entityUpper = String(req.entityContext || "").trim().toUpperCase();
    let assignedApproverSql: number | null = null;
    if (stepsJsonIncludesApproverRole(requestStepsJson)) {
      const aid = Number(assigned_approver_id);
      if (!Number.isFinite(aid) || aid <= 0) {
        return res.status(400).json({ error: "Select an approver for this request (assigned_approver_id)." });
      }
      if (aid === Number(req.user.id)) {
        return res.status(400).json({ error: "You cannot assign yourself as the approver." });
      }
      const approverRow = await fetchSqlUsersettingById(sqlPool, aid);
      if (!approverRow) return res.status(400).json({ error: "Chosen approver was not found." });
      if (!userRowHasApproverRole(approverRow) || !userRowHasEntityAccess(approverRow, entityUpper)) {
        return res.status(400).json({
          error: "Chosen approver must have approver role access for this entity.",
        });
      }
      if (!userDepartmentsMatch(approverRow.department, req.user.department)) {
        return res.status(400).json({ error: "Chosen approver must be in one of your departments." });
      }
      assignedApproverSql = aid;
    }

    const taxVal = tax_rate !== undefined ? tax_rate : 0.18;
    const discVal = discount_rate !== undefined ? Number(discount_rate) : 0;
    const curVal = currencyTrimmed;
    const lineJson = JSON.stringify(line_items || []);
    const detailsForDb = isPR || isPO || isSR ? "" : String(details ?? "");

    let requestId: number;
    try {
      const ins = await sqlPool
        .request()
        .input("template_id", sql.Int, Number(template_id))
        .input("requester_id", sql.Int, Number(req.user.id))
        .input("requester_username_snapshot", sql.NVarChar, req.user.username || "")
        .input("requester_name", sql.NVarChar, req.user.username || "")
        .input("template_name_snapshot", sql.NVarChar, template.name || "")
        .input("department", sql.NVarChar, req.user.department)
        .input("title", sql.NVarChar, title)
        .input("details", sql.NVarChar, detailsForDb)
        .input("line_items", sql.NVarChar(sql.MAX), lineJson)
        .input("requester_signature", sql.NVarChar(sql.MAX), requesterSigStored)
        .input("requester_signed_flag", sql.Bit, requesterSigned ? 1 : 0)
        .input("tax_rate", sql.Decimal(10, 6), taxVal)
        .input("discount_rate", sql.Decimal(10, 6), discVal)
        .input("currency", sql.NVarChar, curVal)
        .input("cost_center", sql.NVarChar, cost_center || "")
        .input("section", sql.NVarChar, sectionTrimmed)
        .input("suggested_supplier", sql.NVarChar, suggestedSupplierSql)
        .input("entity", sql.NVarChar, req.entityContext)
        .input("formatted_id", sql.NVarChar, formatted_id)
        .input("request_steps", sql.NVarChar(sql.MAX), requestStepsJson)
        .input("assigned_approver_id", sql.Int, assignedApproverSql)
        .query(`
            INSERT INTO workflow_requests (
              template_id, requester_id, requester_username_snapshot, requester_name, template_name_snapshot,
              department, title, details, line_items, requester_signature, requester_signed_at, tax_rate, discount_rate, currency, cost_center, section,
              suggested_supplier, entity, formatted_id, request_steps, assigned_approver_id
            ) VALUES (
              @template_id, @requester_id, @requester_username_snapshot, @requester_name, @template_name_snapshot,
              @department, @title, @details, @line_items, @requester_signature,
              CASE WHEN @requester_signed_flag = 1 THEN SYSUTCDATETIME() ELSE NULL END,
              @tax_rate, @discount_rate, @currency, @cost_center, @section,
              @suggested_supplier, @entity, @formatted_id, @request_steps, @assigned_approver_id
            );
            SELECT CAST(SCOPE_IDENTITY() AS INT) AS id;
          `);
      requestId = ins.recordset?.[0]?.id;
      if (attachments && Array.isArray(attachments)) {
        for (const att of attachments) {
          const buf = decodeAttachmentPayload(att.data || "");
          if (!buf.length) continue;
          let rel: string | null = null;
          try {
            const saved = saveRequestAttachmentFile(
              req.entityContext,
              req.user.department,
              formatted_id,
              requestId,
              String(att.name || "file"),
              buf,
              String(att.type || "")
            );
            rel = saved.relativePath;
            att.name = saved.storedFileName;
            att.type = saved.mimeType;
          } catch (e: any) {
            console.error("Attachment save failed:", e?.message || e);
            return res.status(400).json({ error: e?.message || "Failed to save attachment file" });
          }
          await sqlPool
            .request()
            .input("request_id", sql.Int, requestId)
            .input("file_name", sql.NVarChar, att.name)
            .input("file_type", sql.NVarChar, att.type || "")
            .input("file_data", sql.NVarChar(sql.MAX), null)
            .input("file_path", sql.NVarChar, rel)
            .query(
              "INSERT INTO request_attachments (request_id, file_name, file_type, file_data, file_path) VALUES (@request_id, @file_name, @file_type, @file_data, @file_path)"
            );
        }
      }
      approvalEventLog(
        `REQUEST_CREATE ${formatActor(req)} request_id=${requestId} template_id=${Number(template_id)} title=${approvalLogSanitize(String(title ?? ""), 240)} formatted_id=${approvalLogSanitize(String(formatted_id ?? ""), 64)}`
      );
      return res.json({ id: requestId });
    } catch (e: any) {
      console.error("SQL POST /api/workflow-requests:", e?.message || e);
      return res.status(400).json({ error: e?.message || "Failed to create request" });
    }
  });

  app.get(
    "/api/workflow-requests/:id/attachments/:attachmentId/file",
    authenticate,
    requireEntityContext,
    async (req: any, res) => {
      const requestId = Number(req.params.id);
      const attachmentId = Number(req.params.attachmentId);
      const rr = await sqlPool.request().input("id", sql.Int, requestId).query(`
        SELECT r.*, w.category AS template_category
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE r.id = @id
      `);
      const requestRow = rr.recordset?.[0];
      if (!assertWorkflowRequestAccess(req, res, requestRow, requestRow?.template_category)) return;
      const rowRs = await sqlPool
        .request()
        .input("request_id", sql.Int, requestId)
        .input("aid", sql.Int, attachmentId)
        .query(
          "SELECT TOP 1 file_name, file_type, file_data, file_path FROM request_attachments WHERE id = @aid AND request_id = @request_id"
        );
      const row: any = rowRs.recordset?.[0];
      if (!row) return res.status(404).json({ error: "Attachment not found" });
      const disp = String(row.file_name || "download").replace(/[^\w.\- ()]+/g, "_");
      try {
        if (row.file_path) {
          const full = resolveStoredPath(row.file_path);
          if (!fs.existsSync(full)) return res.status(404).json({ error: "File missing on server" });
          res.setHeader("Content-Type", row.file_type || "application/octet-stream");
          res.setHeader("X-Content-Type-Options", "nosniff");
          res.setHeader("Content-Disposition", `attachment; filename="${disp}"`);
          return res.sendFile(path.resolve(full));
        }
        if (row.file_data) {
          const buf = decodeAttachmentPayload(row.file_data);
          res.setHeader("Content-Type", row.file_type || "application/octet-stream");
          res.setHeader("X-Content-Type-Options", "nosniff");
          res.setHeader("Content-Disposition", `attachment; filename="${disp}"`);
          return res.send(buf);
        }
        return res.status(404).json({ error: "No file content" });
      } catch (e: any) {
        console.error("Attachment file serve:", e?.message || e);
        return res.status(500).json({ error: "Failed to read file" });
      }
    }
  );

  app.get("/api/workflow-requests/:id/attachments", authenticate, requireEntityContext, async (req: any, res) => {
    const rr = await sqlPool.request().input("id", sql.Int, Number(req.params.id)).query(`
        SELECT r.*, w.category AS template_category
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE r.id = @id
      `);
    const requestRow = rr.recordset?.[0];
    if (!assertWorkflowRequestAccess(req, res, requestRow, requestRow?.template_category)) return;
    const att = await sqlPool
      .request()
      .input("request_id", sql.Int, Number(req.params.id))
      .query(
        "SELECT id, file_name, file_type, file_data, file_path FROM request_attachments WHERE request_id = @request_id"
      );
    const rid = Number(req.params.id);
    const list = (att.recordset || []).map((a: any) => {
      const hasPath = !!(a.file_path && String(a.file_path).trim());
      const hasData = !!(a.file_data && String(a.file_data).trim());
      const file_url =
        hasPath || hasData ? `/api/workflow-requests/${rid}/attachments/${a.id}/file` : null;
      return {
        id: a.id,
        file_name: a.file_name,
        file_type: a.file_type,
        file_path: a.file_path,
        file_data: null,
        file_url,
      };
    });
    return res.json(list);
  });

  app.get("/api/workflow-requests/:id/approvals", authenticate, requireEntityContext, async (req: any, res) => {
    const rr = await sqlPool.request().input("id", sql.Int, Number(req.params.id)).query(`
        SELECT r.*, w.category AS template_category
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE r.id = @id
      `);
    const requestRow = rr.recordset?.[0];
    if (!assertWorkflowRequestAccess(req, res, requestRow, requestRow?.template_category)) return;
    const appr = await sqlPool.request().input("request_id", sql.Int, Number(req.params.id)).query(`
        SELECT a.*, u.username AS approver_name, u.designation AS approver_designation
        FROM request_approvals a
        INNER JOIN ${userJoinSql} u ON a.approver_id = u.id
        WHERE a.request_id = @request_id
        ORDER BY a.created_at ASC
      `);
    return res.json(appr.recordset || []);
  });

  app.get("/api/procurement/requests", authenticate, hasPermission("view_procurement_center"), async (req: any, res) => {
    const { status, department, search, entity } = req.query;
    const crossDept = isAdminUser(req) || isPurchasingUser(req) || isSomUser(req);

    try {
      const reqSql = sqlPool.request();
      const requestedEntity =
        String(entity ?? req.headers["x-entity"] ?? req.headers["X-Entity"] ?? "").trim();
      const allowedEntities = parseUserEntities(req.user).map((e) => e.trim()).filter(Boolean);
      const allowedSetUpper = new Set(allowedEntities.map((e) => e.toUpperCase()));
      const canViewAllEntities = isAdminUser(req) || isSomUser(req) || isDirectorUser(req) || isPurchasingUser(req);
      let entityClause = "";
      if (requestedEntity) {
        if (!canViewAllEntities && !allowedSetUpper.has(requestedEntity.toUpperCase())) {
          return res.status(403).json({ error: "You do not have access to this entity" });
        }
        entityClause = " AND r.entity = @entity";
        reqSql.input("entity", sql.NVarChar, requestedEntity);
      } else if (!canViewAllEntities) {
        if (allowedEntities.length === 0) return res.json([]);
        const params: string[] = [];
        allowedEntities.forEach((ent, idx) => {
          const key = `ent_${idx}`;
          reqSql.input(key, sql.NVarChar, ent);
          params.push(`@${key}`);
        });
        entityClause = ` AND r.entity IN (${params.join(", ")})`;
      }
      let q = `
      SELECT r.*,
        COALESCE(NULLIF(LTRIM(RTRIM(CAST(r.requester_name AS NVARCHAR(255)))), N''), u.username) AS computed_requester_name,
        u.designation AS requester_designation,
        w.name AS template_name, COALESCE(r.request_steps, w.steps) AS template_steps,
        w.table_columns AS table_columns, w.attachments_required AS attachments_required, w.category AS category
      FROM workflow_requests r
      INNER JOIN ${userJoinSql} u ON r.requester_id = u.id
      INNER JOIN workflows w ON r.template_id = w.id
      WHERE w.category = N'procurement'${entityClause}
    `;
      if (!crossDept) {
        q += " AND r.department = @requesterDepartment";
        reqSql.input("requesterDepartment", sql.NVarChar, req.user.department);
      }
      if (status) {
        q += " AND r.status = @status";
        reqSql.input("status", sql.NVarChar, String(status));
      }
      if (crossDept && department) {
        q += " AND r.department = @filterDepartment";
        reqSql.input("filterDepartment", sql.NVarChar, String(department));
      }
      if (search) {
        q += " AND (CAST(r.id AS NVARCHAR(50)) LIKE @search OR r.title LIKE @search)";
        reqSql.input("search", sql.NVarChar, `%${search}%`);
      }
      q += " ORDER BY r.created_at DESC";
      const rs = await reqSql.query(q);
      const requests = rs.recordset || [];
      const mapped = await Promise.all(
        requests.map(async (r: any) => {
          const { computed_requester_name, ...row } = r;
          const approvalsRs = await sqlPool
            .request()
            .input("rid", sql.Int, Number(r.id))
            .query(`
            SELECT a.*, u.username AS approver_name, u.designation AS approver_designation
            FROM request_approvals a
            INNER JOIN ${userJoinSql} u ON a.approver_id = u.id
            WHERE a.request_id = @rid
            ORDER BY a.created_at ASC
          `);
          return {
            ...row,
            requester_name: computed_requester_name ?? row.requester_name,
            template_steps: JSON.parse(row.template_steps),
            table_columns: row.table_columns ? JSON.parse(row.table_columns) : [],
            line_items: row.line_items ? JSON.parse(row.line_items) : [],
            attachments_required: !!row.attachments_required,
            approvals: approvalsRs.recordset || [],
          };
        })
      );
      return res.json(mapped);
    } catch (e: any) {
      console.error("SQL procurement/requests:", e?.message || e);
      return res.status(500).json({ error: "Failed to load procurement requests" });
    }
  });

  app.patch("/api/workflow-requests/:id", authenticate, requireEntityContext, async (req: any, res) => {
    const {
      title,
      details,
      line_items,
      tax_rate,
      discount_rate,
      currency,
      cost_center,
      section,
      suggested_supplier,
      attachment_keep_ids,
      attachments_add,
    } =
      req.body;
    const requestId = req.params.id;

    const rq = await sqlPool.request().input("id", sql.Int, Number(requestId)).query(`
        SELECT r.*, COALESCE(r.request_steps, w.steps) AS template_steps, w.name AS template_name, w.category AS template_category
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE r.id = @id
      `);
    const request: any = rq.recordset?.[0];

    if (!assertWorkflowRequestAccess(req, res, request, request?.template_category)) return;

    const isPR = isPRName(request.template_name) && request.template_category === "procurement";
    const isPurchasing = req.user.roles && req.user.roles.some((r: string) => r.toLowerCase() === "purchasing");
    const userIsRequester = req.user.id === request.requester_id;
    const userHasEditPermission =
      req.user.permissions && (req.user.permissions.includes("edit_requests") || req.user.permissions.includes("admin"));
    const allowRequesterRejectedPrEdit = request.status === "rejected" && isPR && userIsRequester;
    const allowAdminRejectedPrEdit = request.status === "rejected" && isPR && !!req.user.permissions?.includes("admin");

    if (request.status !== "pending") {
      if (request.status === "approved" && isPR && isPurchasing) {
        // Allow purchasing to edit approved PRs
      } else if (allowRequesterRejectedPrEdit || allowAdminRejectedPrEdit) {
        // Allow requester/admin to edit rejected PR before resubmission
      } else {
        return res.status(400).json({ error: "Only pending requests can be edited" });
      }
    }

    const steps = JSON.parse(request.template_steps);
    const currentStep = request.status === "pending" ? steps[request.current_step_index] : null;

    const userHasEditPermissionResolved = userHasEditPermission;
    const roleMatch =
      currentStep &&
      req.user.roles &&
      req.user.roles.some((r: string) => r.toLowerCase() === currentStep.approverRole.toLowerCase());
    const userIsCurrentApprover =
      currentStep &&
      ((roleMatch && userDepartmentsMatch(req.user.department, request.department)) ||
        (currentStep.approverRole.toLowerCase() === "som" && isSomUser(req)));
    const userIsRequesterResolved = userIsRequester;

    if (
      !userHasEditPermissionResolved &&
      !userIsRequesterResolved &&
      !(request.status === "approved" && isPR && isPurchasing) &&
      !allowRequesterRejectedPrEdit &&
      !allowAdminRejectedPrEdit
    ) {
      return res.status(403).json({ error: "You do not have permission to edit this request" });
    }

    if (
      request.status === "pending" &&
      userHasEditPermissionResolved &&
      !userIsCurrentApprover &&
      !userIsRequesterResolved &&
      !req.user.permissions.includes("admin")
    ) {
      return res.status(403).json({ error: "Only the current approver or the requester can edit this request" });
    }

    const rid = Number(requestId);
    const isPO = isPOName(request.template_name) && request.template_category === "procurement";
    const isProcurementPRorPO =
      request.template_category === "procurement" &&
      (isPRName(request.template_name) || isPOName(request.template_name) || isSRName(request.template_name));
    try {
      if (title) {
        await sqlPool.request().input("title", sql.NVarChar, title).input("id", sql.Int, rid).query("UPDATE workflow_requests SET title = @title WHERE id = @id");
      }
      if (isProcurementPRorPO) {
        await sqlPool.request().input("id", sql.Int, rid).query("UPDATE workflow_requests SET details = N'' WHERE id = @id");
      } else if (details) {
        await sqlPool.request().input("details", sql.NVarChar, details).input("id", sql.Int, rid).query("UPDATE workflow_requests SET details = @details WHERE id = @id");
      }
      if (line_items) {
        await sqlPool
          .request()
          .input("line_items", sql.NVarChar(sql.MAX), JSON.stringify(line_items))
          .input("id", sql.Int, rid)
          .query("UPDATE workflow_requests SET line_items = @line_items WHERE id = @id");
      }
      if (tax_rate !== undefined) {
        await sqlPool.request().input("tax_rate", sql.Decimal(10, 6), tax_rate).input("id", sql.Int, rid).query("UPDATE workflow_requests SET tax_rate = @tax_rate WHERE id = @id");
      }
      if (discount_rate !== undefined) {
        await sqlPool
          .request()
          .input("discount_rate", sql.Decimal(10, 6), discount_rate)
          .input("id", sql.Int, rid)
          .query("UPDATE workflow_requests SET discount_rate = @discount_rate WHERE id = @id");
      }
      if (currency !== undefined) {
        const curPatch = String(currency).trim();
        await sqlPool.request().input("currency", sql.NVarChar, curPatch).input("id", sql.Int, rid).query("UPDATE workflow_requests SET currency = @currency WHERE id = @id");
      }
      if (cost_center) {
        await sqlPool.request().input("cost_center", sql.NVarChar, cost_center).input("id", sql.Int, rid).query("UPDATE workflow_requests SET cost_center = @cost_center WHERE id = @id");
      }
      if (section !== undefined) {
        const sectionPatchRaw = String(section ?? "").trim();
        const sectionPatch = sectionPatchRaw.toUpperCase() === "NA" ? "" : sectionPatchRaw;
        await sqlPool
          .request()
          .input("section", sql.NVarChar, sectionPatch)
          .input("id", sql.Int, rid)
          .query("UPDATE workflow_requests SET section = @section WHERE id = @id");
      }
      if (suggested_supplier !== undefined && isPR) {
        const ss = String(suggested_supplier ?? "").trim();
        if (!ss) {
          return res.status(400).json({ error: "Suggested supplier cannot be empty for a purchase request" });
        }
        await sqlPool
          .request()
          .input("suggested_supplier", sql.NVarChar, ss)
          .input("id", sql.Int, rid)
          .query("UPDATE workflow_requests SET suggested_supplier = @suggested_supplier WHERE id = @id");
      }
      const attachmentKeepIds =
        attachment_keep_ids !== undefined && Array.isArray(attachment_keep_ids)
          ? attachment_keep_ids
              .map((v: any) => Number(v))
              .filter((v: number) => Number.isInteger(v) && v > 0)
          : null;
      const attachmentAdds = Array.isArray(attachments_add) ? attachments_add : null;
      if (attachmentKeepIds !== null || attachmentAdds !== null) {
        const existing = await sqlPool
          .request()
          .input("request_id", sql.Int, rid)
          .query("SELECT id, file_path FROM request_attachments WHERE request_id = @request_id");
        if (attachmentKeepIds !== null) {
          const keepSet = new Set<number>(attachmentKeepIds);
          for (const row of existing.recordset || []) {
            const aid = Number((row as any).id);
            if (keepSet.has(aid)) continue;
            tryUnlinkStoredFile((row as any).file_path);
            await sqlPool
              .request()
              .input("id", sql.Int, aid)
              .input("request_id", sql.Int, rid)
              .query("DELETE FROM request_attachments WHERE id = @id AND request_id = @request_id");
          }
        }
        if (attachmentAdds) {
          for (const att of attachmentAdds) {
            const buf = decodeAttachmentPayload(att?.data || "");
            let rel: string | null = null;
            try {
              const saved = saveRequestAttachmentFile(
                req.entityContext,
                request.department,
                request.formatted_id,
                rid,
                String(att?.name || "file"),
                buf,
                String(att?.type || "")
              );
              rel = saved.relativePath;
              att.name = saved.storedFileName;
              att.type = saved.mimeType;
            } catch (e: any) {
              return res.status(400).json({ error: e?.message || "Failed to save attachment file" });
            }
            await sqlPool
              .request()
              .input("request_id", sql.Int, rid)
              .input("file_name", sql.NVarChar, String(att?.name || "file"))
              .input("file_type", sql.NVarChar, String(att?.type || "application/octet-stream"))
              .input("file_data", sql.NVarChar(sql.MAX), null)
              .input("file_path", sql.NVarChar, rel)
              .query(
                "INSERT INTO request_attachments (request_id, file_name, file_type, file_data, file_path) VALUES (@request_id, @file_name, @file_type, @file_data, @file_path)"
              );
          }
        }
      }
      if (request.status === "pending" && isPO && request.current_step_index === 0) {
        if (line_items !== undefined || tax_rate !== undefined || currency !== undefined || discount_rate !== undefined) {
          const li = line_items !== undefined ? line_items : JSON.parse(request.line_items || "[]");
          const tr = tax_rate !== undefined ? tax_rate : request.tax_rate;
          const dr = discount_rate !== undefined ? discount_rate : request.discount_rate;
          const cur =
            currency !== undefined ? String(currency).trim() : String(request.currency ?? "").trim();
          const newSteps = buildPoRequestSteps(computeRequestTotalMyr(li, tr, cur, dr));
          await sqlPool
            .request()
            .input("request_steps", sql.NVarChar(sql.MAX), JSON.stringify(newSteps))
            .input("id", sql.Int, rid)
            .query("UPDATE workflow_requests SET request_steps = @request_steps WHERE id = @id");
        }
      }
      approvalEventLog(
        `REQUEST_UPDATE ${formatActor(req)} request_id=${rid} title=${approvalLogSanitize(String(request.title || ""), 200)}`
      );
      return res.json({ success: true });
    } catch (err: any) {
      return res.status(400).json({ error: err.message });
    }
  });

  app.post("/api/workflow-requests/:id/resubmit", authenticate, requireEntityContext, async (req: any, res) => {
    const requestId = req.params.id;
    const rq = await sqlPool.request().input("id", sql.Int, Number(requestId)).query(`
        SELECT r.*, w.name AS template_name, w.category AS template_category
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE r.id = @id
      `);
    const request: any = rq.recordset?.[0];

    if (!assertWorkflowRequestAccess(req, res, request, request?.template_category)) return;
    if (request.status !== "rejected") return res.status(400).json({ error: "Only rejected requests can be resubmitted" });
    if (request.requester_id !== req.user.id) return res.status(403).json({ error: "Only the requester can resubmit this request" });

    const rid = Number(requestId);
    // Clear prior approval chain so a resubmission requires fresh signatures/approvals.
    // Keep purchasing_final decisions for audit trail (they don't map to a workflow step index).
    try {
      await sqlPool
        .request()
        .input("request_id", sql.Int, rid)
        .query(
          "DELETE FROM request_approvals WHERE request_id = @request_id AND ISNULL(approver_role_snapshot, N'') <> N'purchasing_final'"
        );
    } catch (e: any) {
      console.warn("SQL resubmit clear approvals:", e?.message || e);
    }
    if (isPOName(request.template_name) && request.template_category === "procurement") {
      let lineItems: any[] = [];
      try {
        lineItems = JSON.parse(request.line_items || "[]");
      } catch {
        lineItems = [];
      }
      const poSteps = buildPoRequestSteps(
        computeRequestTotalMyr(lineItems, request.tax_rate, String(request.currency ?? "").trim(), request.discount_rate)
      );
      await sqlPool
        .request()
        .input("request_steps", sql.NVarChar(sql.MAX), JSON.stringify(poSteps))
        .input("id", sql.Int, rid)
        .query("UPDATE workflow_requests SET request_steps = @request_steps WHERE id = @id");
    }

    await sqlPool.request().input("id", sql.Int, rid).query("UPDATE workflow_requests SET status = N'pending', current_step_index = 0 WHERE id = @id");
    approvalEventLog(`REQUEST_RESUBMIT ${formatActor(req)} request_id=${rid} formatted_id=${approvalLogSanitize(String(request.formatted_id || ""), 64)}`);
    res.json({ success: true });
  });

  /**
   * Requester cancellation before any approval is granted:
   * - only requester can cancel
   * - request must still be pending
   * - cannot cancel once any approver has approved
   * The request remains in the system with status = cancelled for audit/history.
   */
  app.post("/api/workflow-requests/:id/cancel", authenticate, requireEntityContext, async (req: any, res) => {
    const requestId = Number(req.params.id);
    const comment = String(req.body?.comment || "").trim();
    if (!Number.isFinite(requestId) || requestId <= 0) return res.status(400).json({ error: "Invalid request id" });

    const rq = await sqlPool.request().input("id", sql.Int, requestId).query(`
      SELECT r.*, w.category AS template_category
      FROM workflow_requests r
      INNER JOIN workflows w ON r.template_id = w.id
      WHERE r.id = @id
    `);
    const request: any = rq.recordset?.[0];

    if (!assertWorkflowRequestAccess(req, res, request, request?.template_category)) return;
    if (!request) return res.status(404).json({ error: "Request not found" });
    if (request.requester_id !== req.user.id) return res.status(403).json({ error: "Only the requester can cancel this request" });
    if (String(request.status || "").trim().toLowerCase() !== "pending") {
      return res.status(400).json({ error: "Only pending requests can be cancelled" });
    }

    const approvedRs = await sqlPool
      .request()
      .input("request_id", sql.Int, requestId)
      .query(
        "SELECT COUNT(1) AS cnt FROM request_approvals WHERE request_id = @request_id AND LOWER(LTRIM(RTRIM(CAST(status AS NVARCHAR(50))))) = N'approved'"
      );
    const approvedCount = Number(approvedRs.recordset?.[0]?.cnt || 0);
    if (approvedCount > 0) {
      return res.status(400).json({ error: "This request already has approver approval and can no longer be cancelled by requester" });
    }

    try {
      await sqlPool
        .request()
        .input("id", sql.Int, requestId)
        .query("UPDATE workflow_requests SET status = N'cancelled' WHERE id = @id");

      await sqlPool
        .request()
        .input("request_id", sql.Int, requestId)
        .input("step_index", sql.Int, Number(request.current_step_index ?? 0))
        .input("approver_id", sql.Int, Number(req.user.id))
        .input("approver_username", sql.NVarChar, String(req.user.username || ""))
        .input("approver_role_snapshot", sql.NVarChar, "requester_cancel")
        .input("request_title_snapshot", sql.NVarChar, request.title || "")
        .input("request_formatted_id_snapshot", sql.NVarChar, request.formatted_id || null)
        .input("status", sql.NVarChar, "cancelled")
        .input("comment", sql.NVarChar, comment)
        .input("approver_signature", sql.NVarChar(sql.MAX), null)
        .query(`
          INSERT INTO request_approvals (
            request_id, step_index, approver_id, approver_username, approver_role_snapshot,
            request_title_snapshot, request_formatted_id_snapshot, status, comment, approver_signature
          ) VALUES (
            @request_id, @step_index, @approver_id, @approver_username, @approver_role_snapshot,
            @request_title_snapshot, @request_formatted_id_snapshot, @status, @comment, @approver_signature
          )
        `);
    } catch (e: any) {
      console.error("SQL requester-cancel:", e?.message || e);
      return res.status(400).json({ error: e?.message || "Cancel failed" });
    }

    approvalEventLog(
      `REQUEST_CANCEL_BY_REQUESTER ${formatActor(req)} request_id=${requestId} formatted_id=${approvalLogSanitize(String(request.formatted_id || ""), 64)} title=${approvalLogSanitize(String(request.title || ""), 200)}`
    );
    return res.json({ success: true });
  });

  // Admin-only: delete a workflow request permanently (request + attachments + approvals)
  app.delete("/api/workflow-requests/:id", authenticate, requireEntityContext, async (req: any, res) => {
    return res.status(403).json({ error: "Delete operations are disabled in this system." });
  });

  app.post("/api/workflow-requests/:id/convert-to-po", authenticate, requireEntityContext, async (req: any, res) => {
    const requestId = req.params.id;
    const isPurchasing = req.user.roles && req.user.roles.some((r: string) => r.toLowerCase() === "purchasing");
    const isAdmin = req.user.permissions && req.user.permissions.includes("admin");
    if (!isPurchasing && !isAdmin) {
      return res.status(403).json({ error: "Only purchasing staff can convert PR to PO" });
    }

    const rq = await sqlPool.request().input("id", sql.Int, Number(requestId)).query(`
        SELECT r.*, w.name AS template_name, w.category AS template_category
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE r.id = @id
      `);
    const request: any = rq.recordset?.[0];

    if (!assertWorkflowRequestAccess(req, res, request, request?.template_category)) return;
    if (request.status !== "approved") return res.status(400).json({ error: "Only approved requests can be converted" });
    if (!isPRName(request.template_name) || request.template_category !== "procurement") {
      return res.status(400).json({ error: "Only Purchase Requisitions can be converted to PO" });
    }

    const existingPoId = request.converted_po_request_id != null ? Number(request.converted_po_request_id) : null;
    if (existingPoId != null && !Number.isNaN(existingPoId) && existingPoId > 0) {
      return res.status(400).json({
        error: "This PR has already been converted to a PO. Each purchase requisition can only create one PO.",
      });
    }

    const pt = await sqlPool.request().query(`
        SELECT TOP 1 id FROM workflows
        WHERE category = N'procurement' AND (name LIKE N'%Purchase Order%' OR name LIKE N'%PO%') AND status = N'approved'
      `);
    const poTemplate: any = pt.recordset?.[0];
    if (!poTemplate) {
      return res.status(400).json({ error: "No approved Purchase Order template found" });
    }

    const poNumberRaw = req.body.po_number != null ? String(req.body.po_number).trim() : "";
    if (!poNumberRaw) {
      return res.status(400).json({
        error:
          "PO number is required. Enter the official purchase order reference (must be unique for this entity).",
      });
    }

    const dupRs = await sqlPool
      .request()
      .input("entity", sql.NVarChar, String(request.entity || "").trim())
      .input("fid", sql.NVarChar, poNumberRaw)
      .query(`
        SELECT TOP 1 r.id,
          r.status,
          r.line_items,
          r.tax_rate,
          r.discount_rate,
          r.currency,
          r.department,
          r.template_id,
          w.name AS template_name,
          w.category AS template_category
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE entity = @entity
          AND formatted_id IS NOT NULL
          AND LTRIM(RTRIM(formatted_id)) <> N''
          AND LOWER(LTRIM(RTRIM(formatted_id))) = LOWER(LTRIM(RTRIM(@fid)))
      `);
    const existingSameNumber: any = dupRs.recordset?.[0];

    const formatted_id = poNumberRaw;

    let lineItems: any[] = [];
    try {
      const prSup = String(request.suggested_supplier || "").trim();
      lineItems = JSON.parse(request.line_items).map((item: any) => ({
        ...item,
        "Final Supplier":
          prSup || String(item["Suggested Supplier"] || item["Final Supplier"] || "").trim() || "",
      }));
    } catch {
      lineItems = [];
    }

    const totalMyr = computeRequestTotalMyr(lineItems, request.tax_rate, String(request.currency ?? "").trim(), request.discount_rate);
    const poRequestStepsJson = JSON.stringify(buildPoRequestSteps(totalMyr));
    const lineJson = JSON.stringify(lineItems);

    let newRequestId: number;
    try {
      if (existingSameNumber) {
        const isExistingPo =
          String(existingSameNumber.template_category || "").toLowerCase() === "procurement" &&
          isPOName(String(existingSameNumber.template_name || ""));
        if (!isExistingPo) {
          return res.status(400).json({
            error:
              "This document number already exists for this entity and is not a PO. Use a different PO number.",
          });
        }
        if (String(existingSameNumber.status || "").toLowerCase() !== "pending") {
          return res.status(400).json({
            error:
              "This PO number already exists, but that PO is no longer pending. Use a new PO number to convert this PR.",
          });
        }

        newRequestId = Number(existingSameNumber.id);
        let existingPoLines: any[] = [];
        try {
          existingPoLines = JSON.parse(existingSameNumber.line_items || "[]");
          if (!Array.isArray(existingPoLines)) existingPoLines = [];
        } catch {
          existingPoLines = [];
        }
        const mergedLines = [...existingPoLines, ...lineItems];
        const mergedLineJson = JSON.stringify(mergedLines);
        const totalMyrMerged = computeRequestTotalMyr(
          mergedLines,
          existingSameNumber.tax_rate,
          String(existingSameNumber.currency ?? "").trim(),
          existingSameNumber.discount_rate
        );
        const mergedStepsJson = JSON.stringify(buildPoRequestSteps(totalMyrMerged));

        await sqlPool
          .request()
          .input("id", sql.Int, newRequestId)
          .input("line_items", sql.NVarChar(sql.MAX), mergedLineJson)
          .input("request_steps", sql.NVarChar(sql.MAX), mergedStepsJson)
          .query(`
            UPDATE workflow_requests
            SET line_items = @line_items,
                request_steps = @request_steps
            WHERE id = @id;
          `);

        const linkRs = await sqlPool
          .request()
          .input("prId", sql.Int, Number(requestId))
          .input("poId", sql.Int, Number(newRequestId))
          .query(`
            UPDATE workflow_requests
            SET converted_po_request_id = @poId
            WHERE id = @prId AND (converted_po_request_id IS NULL OR converted_po_request_id = 0);
          `);
        const rowsAff = linkRs.rowsAffected as number[] | undefined;
        const linkedCount = rowsAff && rowsAff.length ? Number(rowsAff[0]) : 0;
        if (!linkedCount) {
          return res.status(409).json({
            error: "This PR was already converted to a PO (concurrent update). Refresh and open the linked PO request.",
          });
        }

        const attRs = await sqlPool
          .request()
          .input("request_id", sql.Int, Number(requestId))
          .query("SELECT file_name, file_type, file_data, file_path FROM request_attachments WHERE request_id = @request_id");
        const ent = String(request.entity || req.entityContext || "default");
        for (const att of attRs.recordset || []) {
          let relOut: string | null = null;
          const copied = copyStoredFileToRequest(
            ent,
            existingSameNumber.department,
            formatted_id,
            newRequestId,
            att.file_path,
            att.file_name
          );
          if (copied) {
            relOut = copied.relativePath;
          } else if (att.file_data) {
            const buf = decodeAttachmentPayload(att.file_data);
            if (buf.length > 0) {
              relOut = saveRequestAttachmentFile(
                ent,
                existingSameNumber.department,
                formatted_id,
                newRequestId,
                String(att.file_name || "file"),
                buf
              ).relativePath;
            }
          }
          await sqlPool
            .request()
            .input("request_id", sql.Int, newRequestId)
            .input("file_name", sql.NVarChar, att.file_name)
            .input("file_type", sql.NVarChar, att.file_type || "")
            .input("file_data", sql.NVarChar(sql.MAX), null)
            .input("file_path", sql.NVarChar, relOut)
            .query(
              "INSERT INTO request_attachments (request_id, file_name, file_type, file_data, file_path) VALUES (@request_id, @file_name, @file_type, @file_data, @file_path)"
            );
        }

        approvalEventLog(
          `REQUEST_CONVERT_PR_TO_PO_APPEND ${formatActor(req)} pr_request_id=${Number(requestId)} po_request_id=${newRequestId} po_formatted_id=${approvalLogSanitize(String(formatted_id), 64)}`
        );
        return res.json({ id: newRequestId, formatted_id, merged_into_existing: true });
      }

      const ins = await sqlPool
        .request()
        .input("template_id", sql.Int, Number(poTemplate.id))
        .input("requester_id", sql.Int, Number(req.user.id))
        .input("requester_username_snapshot", sql.NVarChar, req.user.username || "")
        .input("requester_name", sql.NVarChar, req.user.username || "")
        .input("template_name_snapshot", sql.NVarChar, "Purchase Order")
        .input("department", sql.NVarChar, request.department)
        .input("title", sql.NVarChar, `PO for: ${request.title}`)
        .input("details", sql.NVarChar, "")
        .input("line_items", sql.NVarChar(sql.MAX), lineJson)
        .input("requester_signature", sql.NVarChar(sql.MAX), request.requester_signature ?? null)
        .input("po_requester_signed_flag", sql.Bit, request.requester_signed_at || request.requester_signature ? 1 : 0)
        .input("tax_rate", sql.Decimal(10, 6), request.tax_rate)
        .input("discount_rate", sql.Decimal(10, 6), request.discount_rate !== undefined && request.discount_rate !== null ? request.discount_rate : 0)
        .input("currency", sql.NVarChar, request.currency)
        .input("cost_center", sql.NVarChar, request.cost_center || "")
        .input("section", sql.NVarChar, request.section || "")
        .input("entity", sql.NVarChar, request.entity)
        .input("formatted_id", sql.NVarChar, formatted_id)
        .input("request_steps", sql.NVarChar(sql.MAX), poRequestStepsJson)
        .input("assigned_approver_id", sql.Int, null)
        .query(`
            INSERT INTO workflow_requests (
              template_id, requester_id, requester_username_snapshot, requester_name, template_name_snapshot,
              department, title, details, line_items, requester_signature, requester_signed_at, tax_rate, discount_rate, currency, cost_center, section,
              entity, formatted_id, request_steps, assigned_approver_id
            ) VALUES (
              @template_id, @requester_id, @requester_username_snapshot, @requester_name, @template_name_snapshot,
              @department, @title, @details, @line_items, @requester_signature,
              CASE WHEN @po_requester_signed_flag = 1 THEN SYSUTCDATETIME() ELSE NULL END,
              @tax_rate, @discount_rate, @currency, @cost_center, @section,
              @entity, @formatted_id, @request_steps, @assigned_approver_id
            );
            SELECT CAST(SCOPE_IDENTITY() AS INT) AS id;
          `);
      newRequestId = ins.recordset?.[0]?.id;
      const linkRs = await sqlPool
        .request()
        .input("prId", sql.Int, Number(requestId))
        .input("poId", sql.Int, Number(newRequestId))
        .query(`
          UPDATE workflow_requests
          SET converted_po_request_id = @poId
          WHERE id = @prId AND (converted_po_request_id IS NULL OR converted_po_request_id = 0);
        `);
      const rowsAff = linkRs.rowsAffected as number[] | undefined;
      const linkedCount = rowsAff && rowsAff.length ? Number(rowsAff[0]) : 0;
      if (!linkedCount) {
        await sqlPool.request().input("orphanId", sql.Int, Number(newRequestId)).query(`
          DELETE FROM request_attachments WHERE request_id = @orphanId;
          DELETE FROM workflow_requests WHERE id = @orphanId;
        `);
        return res.status(409).json({
          error: "This PR was already converted to a PO (concurrent update). Refresh and open the existing PO request.",
        });
      }
      const attRs = await sqlPool
        .request()
        .input("request_id", sql.Int, Number(requestId))
        .query("SELECT file_name, file_type, file_data, file_path FROM request_attachments WHERE request_id = @request_id");
      const ent = String(request.entity || req.entityContext || "default");
      for (const att of attRs.recordset || []) {
        let relOut: string | null = null;
        const copied = copyStoredFileToRequest(
          ent,
          request.department,
          formatted_id,
          newRequestId,
          att.file_path,
          att.file_name
        );
        if (copied) {
          relOut = copied.relativePath;
        } else if (att.file_data) {
          const buf = decodeAttachmentPayload(att.file_data);
          if (buf.length > 0) {
            relOut = saveRequestAttachmentFile(
              ent,
              request.department,
              formatted_id,
              newRequestId,
              String(att.file_name || "file"),
              buf
            ).relativePath;
          }
        }
        await sqlPool
          .request()
          .input("request_id", sql.Int, newRequestId)
          .input("file_name", sql.NVarChar, att.file_name)
          .input("file_type", sql.NVarChar, att.file_type || "")
          .input("file_data", sql.NVarChar(sql.MAX), null)
          .input("file_path", sql.NVarChar, relOut)
          .query(
            "INSERT INTO request_attachments (request_id, file_name, file_type, file_data, file_path) VALUES (@request_id, @file_name, @file_type, @file_data, @file_path)"
          );
      }
      approvalEventLog(
        `REQUEST_CONVERT_PR_TO_PO ${formatActor(req)} pr_request_id=${Number(requestId)} po_request_id=${newRequestId} po_formatted_id=${approvalLogSanitize(String(formatted_id), 64)}`
      );
      return res.json({ id: newRequestId, formatted_id, merged_into_existing: false });
    } catch (e: any) {
      console.error("SQL convert-to-po:", e?.message || e);
      return res.status(400).json({ error: e?.message || "Convert failed" });
    }
  });

  app.post("/api/workflow-requests/:id/approve", authenticate, requireEntityContext, async (req: any, res) => {
    const { status, comment, approver_signature } = req.body;
    const requestId = req.params.id;
    const approverSigned = bodyIndicatesApproverSigned(req.body);
    const approverSigStored =
      typeof approver_signature === "string" && approver_signature.trim().length > 0 ? approver_signature.trim() : null;

    const rqAppr = await sqlPool.request().input("id", sql.Int, Number(requestId)).query(`
        SELECT r.*, COALESCE(r.request_steps, w.steps) AS template_steps, w.category AS template_category,
          COALESCE(NULLIF(LTRIM(RTRIM(CAST(r.template_name_snapshot AS NVARCHAR(255)))), N''), w.name) AS template_name_resolved
        FROM workflow_requests r
        INNER JOIN workflows w ON r.template_id = w.id
        WHERE r.id = @id
      `);
    const request: any = rqAppr.recordset?.[0];

    if (!assertWorkflowRequestAccess(req, res, request, request?.template_category)) return;
    if (request.status !== "pending") return res.status(400).json({ error: "Request already processed" });

    const steps = JSON.parse(request.template_steps);
    const currentStep = steps[request.current_step_index];
    const actorRow = await fetchSqlUsersettingById(sqlPool, Number(req.user.id));
    const actorRoles = parseCommaSeparatedList(actorRow?.role ?? req.user.roles);
    const actorDepartment = actorRow?.department ?? req.user.department;

    const userIsAdmin = actorRoles.some((r) => r.toLowerCase() === "admin");
    const userIsDirector =
      actorRoles.some((r) => r.toLowerCase() === "director") &&
      userHasDepartment(actorDepartment, "management");
    const roleLower = String(currentStep?.approverRole || "").toLowerCase();
    const isAssignedApproverStep = request.assigned_approver_id != null && roleLower === "approver";
    const assignedId = Number(request.assigned_approver_id);
    const userIsAssignedApprover =
      isAssignedApproverStep && Number.isFinite(assignedId) && assignedId > 0 && Number(req.user.id) === assignedId;

    const userHasRole = actorRoles.some((r) => r.toLowerCase() === currentStep.approverRole.toLowerCase());
    const userHasApproverRole = actorRoles.some((r) => r.toLowerCase() === "approver");
    const userDeptMatchRequest = userDepartmentsMatch(actorDepartment, request.department);
    const userIsSom =
      actorRoles.some((r) => r.toLowerCase() === "som") &&
      userHasDepartment(actorDepartment, "management");
    const somStep = roleLower === "som";
    let approverStepTotalMyr: number | null = null;
    let currentUserLimitMyr: number | null = null;
    let assignedApproverInsufficient = false;
    if (roleLower === "approver") {
      let lineItemsForLimit: any[] = [];
      try {
        lineItemsForLimit = JSON.parse(request.line_items || "[]");
      } catch {
        lineItemsForLimit = [];
      }
      approverStepTotalMyr = computeRequestTotalMyr(
        lineItemsForLimit,
        request.tax_rate,
        String(request.currency ?? "").trim(),
        request.discount_rate
      );
      const entForLimitCheck = String(request.entity ?? "").trim();
      currentUserLimitMyr = await getUserApprovalLimitMyr(sqlPool, Number(req.user.id), entForLimitCheck);
      if (isAssignedApproverStep && Number.isFinite(assignedId) && assignedId > 0) {
        const assignedApproverLimitMyr = await getUserApprovalLimitMyr(sqlPool, assignedId, entForLimitCheck);
        assignedApproverInsufficient =
          assignedApproverLimitMyr !== null &&
          approverStepTotalMyr > assignedApproverLimitMyr;
      }
    }
    const currentUserCanCoverApproverTotal =
      approverStepTotalMyr !== null &&
      (currentUserLimitMyr === null || approverStepTotalMyr <= currentUserLimitMyr);
    const roleBasedApprovalAllowed = isAssignedApproverStep
      ? (userIsAssignedApprover && userDeptMatchRequest) ||
        (!userIsAssignedApprover &&
          !!userHasApproverRole &&
          userDeptMatchRequest &&
          assignedApproverInsufficient &&
          currentUserCanCoverApproverTotal)
      : userHasRole && userDeptMatchRequest;
    const canApprove =
      userIsAdmin ||
      userIsDirector ||
      (somStep && userIsSom) ||
      roleBasedApprovalAllowed;

    if (!canApprove) {
      return res.status(403).json({ error: "You do not have the required role or department access to approve this step" });
    }

    const tplName = String(request.template_name_resolved || "");
    const n = tplName.toLowerCase();
    const isProcurementSignatureDoc =
      request.template_category === "procurement" &&
      (isPRName(tplName) || isPOName(tplName) || isSRName(tplName) || n.includes("invoice"));
    if (status === "approved" && isProcurementSignatureDoc && !approverSigned) {
      return res.status(400).json({ error: "Signature is required to approve this procurement document" });
    }

    if (status === "approved" && currentStep.approverRole.toLowerCase() === "approver") {
      if (approverStepTotalMyr === null) {
        return res.status(400).json({ error: "Unable to compute approval total for this request." });
      }
      const userLimitMyr =
        currentUserLimitMyr !== null
          ? currentUserLimitMyr
          : await getUserApprovalLimitMyr(sqlPool, Number(req.user.id), String(request.entity ?? "").trim());
      if (userLimitMyr !== null && approverStepTotalMyr > userLimitMyr) {
        return res.status(403).json({
          error: `Approval limit exceeded. Request total is RM ${approverStepTotalMyr.toFixed(2)}, your limit is RM ${userLimitMyr.toFixed(2)} (entity-specific cap if configured).`,
        });
      }
    }

    const rid = Number(requestId);
    const approverDisplay = String(req.user.username || "");
    try {
      await sqlPool
        .request()
        .input("request_id", sql.Int, rid)
        .input("step_index", sql.Int, Number(request.current_step_index))
        .input("approver_id", sql.Int, Number(req.user.id))
        .input("approver_username", sql.NVarChar, approverDisplay)
        .input("approver_role_snapshot", sql.NVarChar, currentStep?.approverRole || "")
        .input("request_title_snapshot", sql.NVarChar, request.title || "")
        .input("request_formatted_id_snapshot", sql.NVarChar, request.formatted_id || null)
        .input("status", sql.NVarChar, status)
        .input("comment", sql.NVarChar, comment || "")
        .input("approver_signature", sql.NVarChar(sql.MAX), approverSigStored)
        .query(`
            INSERT INTO request_approvals (
              request_id, step_index, approver_id, approver_username, approver_role_snapshot,
              request_title_snapshot, request_formatted_id_snapshot, status, comment, approver_signature
            ) VALUES (
              @request_id, @step_index, @approver_id, @approver_username, @approver_role_snapshot,
              @request_title_snapshot, @request_formatted_id_snapshot, @status, @comment, @approver_signature
            )
          `);
      if (status === "rejected") {
        await sqlPool.request().input("id", sql.Int, rid).query("UPDATE workflow_requests SET status = N'rejected' WHERE id = @id");
      } else {
        const nextStepIndex = request.current_step_index + 1;
        const stepRoleLower = String(currentStep?.approverRole || "").toLowerCase();
        if (status === "approved" && stepRoleLower === "checker") {
          await sqlPool
            .request()
            .input("id", sql.Int, rid)
            .input("checker_name", sql.NVarChar, approverDisplay)
            .query("UPDATE workflow_requests SET checked_at = SYSUTCDATETIME(), checker_name = @checker_name WHERE id = @id");
        }
        if (nextStepIndex >= steps.length) {
          await sqlPool
            .request()
            .input("id", sql.Int, rid)
            .input("approver_name", sql.NVarChar, approverDisplay)
            .query(
              "UPDATE workflow_requests SET status = N'approved', approved_at = SYSUTCDATETIME(), approver_name = @approver_name WHERE id = @id"
            );
        } else {
          await sqlPool
            .request()
            .input("next_step_index", sql.Int, nextStepIndex)
            .input("id", sql.Int, rid)
            .query("UPDATE workflow_requests SET current_step_index = @next_step_index WHERE id = @id");
        }
      }
    } catch (e: any) {
      console.error("SQL approve:", e?.message || e);
      return res.status(400).json({ error: e?.message || "Approval failed" });
    }

    approvalEventLog(
      `REQUEST_APPROVAL ${formatActor(req)} request_id=${Number(requestId)} step_index=${request.current_step_index} decision=${approvalLogSanitize(String(status), 20)} formatted_id=${approvalLogSanitize(String(request.formatted_id || ""), 64)} title=${approvalLogSanitize(String(request.title || ""), 200)}`
    );
    return res.json({ success: true });
  });

  /**
   * Purchasing final decision for approved PRs:
   * - rejected: requester may edit + resubmit.
   * - cancelled: terminal stop (cannot resubmit).
   */
  app.post("/api/workflow-requests/:id/purchasing-decision", authenticate, requireEntityContext, async (req: any, res) => {
    const decisionRaw = String(req.body?.decision || "").trim().toLowerCase();
    const comment = String(req.body?.comment || "").trim();
    const requestId = Number(req.params.id);
    if (!Number.isFinite(requestId) || requestId <= 0) return res.status(400).json({ error: "Invalid request id" });
    if (decisionRaw !== "rejected" && decisionRaw !== "cancelled") {
      return res.status(400).json({ error: "decision must be either rejected or cancelled" });
    }
    const isPurchasing = req.user.roles && req.user.roles.some((r: string) => r.toLowerCase() === "purchasing");
    const isAdmin = req.user.permissions && req.user.permissions.includes("admin");
    if (!isPurchasing && !isAdmin) {
      return res.status(403).json({ error: "Only purchasing staff can take final PR decision" });
    }

    const rq = await sqlPool.request().input("id", sql.Int, requestId).query(`
      SELECT r.*, COALESCE(r.request_steps, w.steps) AS template_steps, w.name AS template_name, w.category AS template_category
      FROM workflow_requests r
      INNER JOIN workflows w ON r.template_id = w.id
      WHERE r.id = @id
    `);
    const request: any = rq.recordset?.[0];
    if (!assertWorkflowRequestAccess(req, res, request, request?.template_category)) return;
    if (!request) return res.status(404).json({ error: "Request not found" });
    if (request.status !== "approved") return res.status(400).json({ error: "Only approved PR can be rejected/cancelled by purchasing" });
    if (!(request.template_category === "procurement" && isPRName(String(request.template_name || "")))) {
      return res.status(400).json({ error: "Purchasing final decision is only available for PR" });
    }

    let stepCount = 0;
    try {
      const arr = JSON.parse(request.template_steps || "[]");
      stepCount = Array.isArray(arr) ? arr.length : 0;
    } catch {
      stepCount = 0;
    }

    const actorName = String(req.user.username || "");
    try {
      if (decisionRaw === "rejected") {
        // Remove existing approvals/signatures from the prior "approved" cycle.
        // This ensures the approver must sign again after requester edits + resubmits.
        await sqlPool.request().input("request_id", sql.Int, requestId).query(
          "DELETE FROM request_approvals WHERE request_id = @request_id"
        );
      }
      await sqlPool
        .request()
        .input("request_id", sql.Int, requestId)
        .input("step_index", sql.Int, Math.max(0, stepCount))
        .input("approver_id", sql.Int, Number(req.user.id))
        .input("approver_username", sql.NVarChar, actorName)
        .input("approver_role_snapshot", sql.NVarChar, "purchasing_final")
        .input("request_title_snapshot", sql.NVarChar, request.title || "")
        .input("request_formatted_id_snapshot", sql.NVarChar, request.formatted_id || null)
        .input("status", sql.NVarChar, decisionRaw)
        .input("comment", sql.NVarChar, comment)
        .input("approver_signature", sql.NVarChar(sql.MAX), null)
        .query(`
          INSERT INTO request_approvals (
            request_id, step_index, approver_id, approver_username, approver_role_snapshot,
            request_title_snapshot, request_formatted_id_snapshot, status, comment, approver_signature
          ) VALUES (
            @request_id, @step_index, @approver_id, @approver_username, @approver_role_snapshot,
            @request_title_snapshot, @request_formatted_id_snapshot, @status, @comment, @approver_signature
          )
        `);

      await sqlPool
        .request()
        .input("id", sql.Int, requestId)
        .input("status", sql.NVarChar, decisionRaw)
        .query("UPDATE workflow_requests SET status = @status WHERE id = @id");
    } catch (e: any) {
      console.error("SQL purchasing-decision:", e?.message || e);
      return res.status(400).json({ error: e?.message || "Purchasing final decision failed" });
    }

    approvalEventLog(
      `REQUEST_PURCHASING_DECISION ${formatActor(req)} request_id=${requestId} decision=${approvalLogSanitize(decisionRaw, 20)} formatted_id=${approvalLogSanitize(String(request.formatted_id || ""), 64)} title=${approvalLogSanitize(String(request.title || ""), 200)}`
    );
    return res.json({ success: true });
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: {
        middlewareMode: true,
        host: true,
        hmr: process.env.DISABLE_HMR === 'true'
          ? false
          : (process.env.VITE_HMR_HOST ? { host: process.env.VITE_HMR_HOST } : true),
      },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    const lan = getLanIPv4Addresses();
    const line = "─".repeat(56);
    console.log(`\n${line}`);
    console.log(`  Local:    http://localhost:${PORT}`);
    if (lan.length > 0) {
      console.log("  LAN:");
      for (const ip of lan) {
        console.log(`            http://${ip}:${PORT}`);
      }
      console.log("  (Also shown on the login page under 'Open this app'.)");
      console.log(
        "  Firewall: npm run allow-lan  (Admin PowerShell) if other devices cannot connect."
      );
    } else {
      console.log(
        "  LAN:     (none detected — use ipconfig, or open the login page for hints.)"
      );
    }
    console.log(`${line}\n`);
  });
}

startServer();
