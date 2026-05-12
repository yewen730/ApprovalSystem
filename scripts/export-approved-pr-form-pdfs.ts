/**
 * One-time / maintenance: generate F-PU-003 PR PDFs for already-approved Purchase Requests
 * and save them under the same UNC layout as POST /api/workflow-requests/:id/generated-form-pdf
 * (`entity` / `department` / `formatted_id` / `PR_<SERIAL>.pdf`).
 *
 * Usage (from repo root, with .env containing SQLSERVER_CONNECTION_STRING and optional ATTACHMENTS_STORAGE_PATH):
 *   npx tsx scripts/export-approved-pr-form-pdfs.ts
 *   npx tsx scripts/export-approved-pr-form-pdfs.ts --dry-run
 *   npx tsx scripts/export-approved-pr-form-pdfs.ts --force
 *   npx tsx scripts/export-approved-pr-form-pdfs.ts --id=123
 *   npx tsx scripts/export-approved-pr-form-pdfs.ts --limit=50
 *
 * Requires network/SMB access to ATTACHMENTS_STORAGE_PATH (same machine as SQL is typical).
 */
import "dotenv/config";
import sql from "mssql";
import type { RequestApproval, WorkflowRequest, WorkflowStep } from "../src/types.ts";
import { buildProcurementPrFormPdfDoc } from "../src/procurement/prFormPdfGenerator.ts";
import {
  saveGeneratedProcurementFormPdf,
  tryUnlinkStoredFile,
  validateAttachmentUpload,
} from "../attachmentStorage.ts";

const conn = process.env.SQLSERVER_CONNECTION_STRING || "";
const DEFAULT_USERS_TABLE = "[dbo].[usersetting]";
const rawUsers = (process.env.SQLSERVER_USERS_TABLE || DEFAULT_USERS_TABLE).trim();
const USERS_TABLE =
  /^[\w\[\].]+$/.test(rawUsers) && rawUsers.length < 200 ? rawUsers : DEFAULT_USERS_TABLE;

function argFlag(name: string): boolean {
  return process.argv.includes(`--${name}`);
}
function argValue(name: string): string | undefined {
  const p = `--${name}=`;
  const hit = process.argv.find((a) => a.startsWith(p));
  return hit ? hit.slice(p.length) : undefined;
}

const isSRName = (name: string) => {
  const n = name.toLowerCase();
  return n.includes("stock requisition") || (n.includes("stock") && n.includes("requisition"));
};

const isPRName = (name: string) => {
  const n = name.toLowerCase();
  if (isSRName(name)) return false;
  return n.includes("purchase request") || /\bpr\b/.test(n);
};

type DbRow = Record<string, unknown>;

function parseJson<T>(v: unknown, fallback: T): T {
  if (v == null || v === "") return fallback;
  if (typeof v === "object") return v as T;
  try {
    return JSON.parse(String(v)) as T;
  } catch {
    return fallback;
  }
}

function toWorkflowRequest(row: DbRow, approvals: RequestApproval[]): WorkflowRequest {
  const templateSteps = parseJson<WorkflowStep[]>(
    row.request_steps ?? row.workflow_steps,
    []
  );
  const lineItems = parseJson<unknown[]>(row.line_items, []);
  const storedName = String(row.template_name ?? "").trim();
  const joinName = String(row.workflow_template_name ?? "").trim();
  return {
    id: Number(row.id),
    template_id: Number(row.template_id),
    requester_id: Number(row.requester_id),
    requester_username: row.requester_username != null ? String(row.requester_username) : null,
    requester_name: String(row.requester_name ?? ""),
    department: String(row.department ?? ""),
    template_name: storedName || joinName,
    template_steps: templateSteps,
    title: String(row.title ?? ""),
    details: String(row.details ?? ""),
    entity: row.entity != null ? String(row.entity) : undefined,
    formatted_id: row.formatted_id != null ? String(row.formatted_id) : undefined,
    line_items: lineItems as WorkflowRequest["line_items"],
    tax_rate: row.tax_rate != null ? Number(row.tax_rate) : undefined,
    discount_rate: row.discount_rate != null ? Number(row.discount_rate) : null,
    currency: row.currency != null ? String(row.currency) : undefined,
    cost_center: row.cost_center != null ? String(row.cost_center) : undefined,
    section: row.section != null ? String(row.section) : undefined,
    suggested_supplier: row.suggested_supplier != null ? String(row.suggested_supplier) : null,
    requester_signature: row.requester_signature != null ? String(row.requester_signature) : undefined,
    requester_signed_at: row.requester_signed_at != null ? String(row.requester_signed_at) : undefined,
    requester_designation: row.requester_designation != null ? String(row.requester_designation) : undefined,
    status: String(row.status ?? "pending") as WorkflowRequest["status"],
    current_step_index: Number(row.current_step_index ?? 0),
    created_at: String(row.created_at ?? ""),
    approvals,
  };
}

async function removeExistingPrFormAttachments(
  pool: sql.ConnectionPool,
  requestId: number,
  serialRaw: string
): Promise<void> {
  const kind = "PR";
  const logicalName = `${kind}_${serialRaw}.pdf`;
  const legacyFormName = `${kind}_${serialRaw}_Form.pdf`;
  const exReq = pool.request();
  exReq.input("request_id", sql.Int, requestId);
  exReq.input("fnNew", sql.NVarChar, logicalName);
  exReq.input("fnLegacy", sql.NVarChar, legacyFormName);
  exReq.input("fnLikePattern", sql.NVarChar, `${kind}_${serialRaw}%`);
  const existingRows =
    (
      await exReq.query(`
          SELECT id, file_path, file_name FROM request_attachments
          WHERE request_id = @request_id
            AND (
              file_name = @fnNew
              OR file_name = @fnLegacy
              OR LOWER(LTRIM(RTRIM(file_name))) LIKE LOWER(@fnLikePattern)
            )
        `)
    ).recordset || [];
  for (const ex of existingRows) {
    tryUnlinkStoredFile((ex as DbRow).file_path as string);
    await pool
      .request()
      .input("eid", sql.Int, Number((ex as DbRow).id))
      .query("DELETE FROM request_attachments WHERE id = @eid");
  }
}

async function main() {
  if (!conn) {
    console.error("Set SQLSERVER_CONNECTION_STRING in .env");
    process.exit(1);
  }
  const dryRun = argFlag("dry-run");
  const force = argFlag("force");
  const limitRaw = argValue("limit");
  const idRaw = argValue("id");
  const limit = limitRaw ? Math.max(1, parseInt(limitRaw, 10) || 0) : 0;
  const onlyId = idRaw ? Number(idRaw) : NaN;

  const pool = await sql.connect(conn);
  console.log("[export-pr-pdfs] Connected. dryRun=%s force=%s", dryRun, force);

  const idSql =
    Number.isFinite(onlyId) && onlyId > 0 ? ` AND r.id = ${Math.floor(onlyId)} ` : "";
  const topClause = limit > 0 ? `TOP (${limit})` : "";

  const reqRs = await pool.request().query(`
      SELECT ${topClause}
        r.id,
        r.template_id,
        r.requester_id,
        r.requester_username,
        r.entity,
        r.department,
        r.formatted_id,
        r.status,
        r.line_items,
        r.tax_rate,
        r.discount_rate,
        r.currency,
        r.cost_center,
        r.section,
        r.suggested_supplier,
        r.title,
        r.details,
        COALESCE(rs.signature_data, r.requester_signature) AS requester_signature,
        r.requester_signed_at,
        r.template_name,
        r.current_step_index,
        r.created_at,
        r.request_steps,
        w.steps AS workflow_steps,
        w.name AS workflow_template_name,
        u.designation AS requester_designation
      FROM workflow_requests r
      INNER JOIN workflows w ON w.id = r.template_id
      LEFT JOIN dbo.workflow_request_requester_signatures rs ON rs.request_id = r.id
      INNER JOIN ${USERS_TABLE} u ON u.id = r.requester_id
      WHERE w.category = N'procurement'
        AND LOWER(LTRIM(RTRIM(r.status))) = N'approved'
        ${idSql}
      ORDER BY r.id DESC
    `);

  const rows: DbRow[] = reqRs.recordset || [];
  const prRows = rows.filter((r) => isPRName(String(r.template_name || r.workflow_template_name || "")));
  console.log("[export-pr-pdfs] Candidate rows: %d (PR templates after filter: %d)", rows.length, prRows.length);

  const ids = prRows.map((r) => Number(r.id)).filter((n) => Number.isFinite(n) && n > 0);
  if (ids.length === 0) {
    await pool.close();
    return;
  }

  const approvalsMap = new Map<number, RequestApproval[]>();
  const chunkSize = 200;
  for (let i = 0; i < ids.length; i += chunkSize) {
    const chunk = ids.slice(i, i + chunkSize);
    const placeholders = chunk.map((_, j) => `@p${j}`).join(", ");
    const rq = pool.request();
    chunk.forEach((id, j) => rq.input(`p${j}`, sql.Int, id));
    const ar = await rq.query(`
        SELECT a.id, a.request_id, a.step_index, a.approver_id, a.status, a.comment,
          a.approver_role_snapshot, a.approver_signature, a.signed_by_user_id, a.created_at,
          u.username AS approver_name, u.designation AS approver_designation,
          su.username AS signed_by_name
        FROM request_approvals a
        INNER JOIN ${USERS_TABLE} u ON a.approver_id = u.id
        LEFT JOIN ${USERS_TABLE} su ON a.signed_by_user_id = su.id
        WHERE a.request_id IN (${placeholders})
        ORDER BY a.request_id ASC, a.created_at ASC
      `);
    for (const raw of ar.recordset || []) {
      const rid = Number((raw as DbRow).request_id);
      const list = approvalsMap.get(rid) || [];
      list.push({
        id: Number((raw as DbRow).id),
        request_id: rid,
        step_index: Number((raw as DbRow).step_index),
        approver_id: Number((raw as DbRow).approver_id),
        approver_name: String((raw as DbRow).approver_name ?? ""),
        approver_role_snapshot:
          (raw as DbRow).approver_role_snapshot != null
            ? String((raw as DbRow).approver_role_snapshot)
            : undefined,
        approver_designation:
          (raw as DbRow).approver_designation != null
            ? String((raw as DbRow).approver_designation)
            : undefined,
        status: String((raw as DbRow).status ?? "") as RequestApproval["status"],
        comment: String((raw as DbRow).comment ?? ""),
        approver_signature:
          (raw as DbRow).approver_signature != null ? String((raw as DbRow).approver_signature) : undefined,
        signed_by_user_id:
          (raw as DbRow).signed_by_user_id != null ? Number((raw as DbRow).signed_by_user_id) : null,
        signed_by_name: (raw as DbRow).signed_by_name != null ? String((raw as DbRow).signed_by_name) : null,
        created_at: String((raw as DbRow).created_at ?? ""),
      });
      approvalsMap.set(rid, list);
    }
  }

  let ok = 0;
  let skipped = 0;
  let failed = 0;

  for (const row of prRows) {
    const requestId = Number(row.id);
    const serialRaw =
      String(row.formatted_id || "")
        .trim()
        .toUpperCase() || `REQUEST_${requestId}`;
    const approvals = approvalsMap.get(requestId) || [];

    const exCheck = await pool
      .request()
      .input("request_id", sql.Int, requestId)
      .input("fnLike", sql.NVarChar, `PR_${serialRaw}%`)
      .query(`
        SELECT TOP 1 id FROM request_attachments
        WHERE request_id = @request_id
          AND LOWER(LTRIM(RTRIM(file_name))) LIKE LOWER(@fnLike)
          AND file_path IS NOT NULL AND LTRIM(RTRIM(CAST(file_path AS NVARCHAR(400)))) <> N''
      `);
    if (exCheck.recordset?.length && !force) {
      console.log("[skip] id=%s already has PR form attachment", requestId);
      skipped += 1;
      continue;
    }

    const wf = toWorkflowRequest(row, approvals);
    try {
      const doc = buildProcurementPrFormPdfDoc(wf);
      const buf = Buffer.from(doc.output("arraybuffer") as ArrayBuffer);
      validateAttachmentUpload(`PR_${serialRaw}.pdf`, "application/pdf", buf);
      if (dryRun) {
        console.log("[dry-run] id=%s serial=%s bytes=%d", requestId, serialRaw, buf.length);
        ok += 1;
        continue;
      }
      await removeExistingPrFormAttachments(pool, requestId, serialRaw);
      const saved = saveGeneratedProcurementFormPdf(
        String(row.entity || "").trim() || "UNKNOWN",
        row.department != null ? String(row.department) : null,
        row.formatted_id != null ? String(row.formatted_id) : null,
        requestId,
        `PR_${serialRaw}.pdf`,
        buf,
        "application/pdf"
      );
      await pool
        .request()
        .input("request_id", sql.Int, requestId)
        .input("file_name", sql.NVarChar, saved.storedFileName)
        .input("file_type", sql.NVarChar, saved.mimeType)
        .input("file_data", sql.NVarChar(sql.MAX), null)
        .input("file_path", sql.NVarChar, saved.storedPath)
        .query(
          "INSERT INTO request_attachments (request_id, file_name, file_type, file_data, file_path) VALUES (@request_id, @file_name, @file_type, @file_data, @file_path)"
        );
      await pool
        .request()
        .input("id", sql.Int, requestId)
        .input("generated_path", sql.NVarChar, saved.storedPath)
        .query("UPDATE workflow_requests SET generated_form_pdf_path = @generated_path WHERE id = @id");
      console.log("[ok] id=%s -> %s", requestId, saved.storedPath);
      ok += 1;
    } catch (e: unknown) {
      failed += 1;
      console.error("[fail] id=%s", requestId, e);
    }
  }

  await pool.close();
  console.log("[export-pr-pdfs] done. ok=%d skipped=%d failed=%d", ok, skipped, failed);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
