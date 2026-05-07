import "dotenv/config";
import sql from "mssql";

const conn = process.env.SQLSERVER_CONNECTION_STRING || "";

type ReqRow = {
  id: number;
  status: string | null;
  template_name: string | null;
  template_category: string | null;
  request_steps: string | null;
  workflow_steps: string | null;
  converted_po_request_id: number | null;
};

type ApprovalRow = {
  request_id: number;
  step_index: number | null;
  status: string | null;
  approver_role_snapshot: string | null;
  created_at: Date | null;
};

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

function parseStepsLen(requestSteps: string | null, workflowSteps: string | null): number {
  try {
    const raw = requestSteps || workflowSteps || "[]";
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr.length : 0;
  } catch {
    return 0;
  }
}

function computeExpectedStatus(req: ReqRow, approvals: ApprovalRow[]): string | null {
  const name = String(req.template_name || "").trim();
  const category = String(req.template_category || "").trim().toLowerCase();
  if (category !== "procurement") return null;
  if (!isPRName(name) && !isSRName(name) && !isPOName(name)) return null;

  const normalizedApprovals = approvals
    .map((a) => ({
      step_index: Number(a.step_index ?? -1),
      status: String(a.status || "").trim().toLowerCase(),
      role: String(a.approver_role_snapshot || "").trim().toLowerCase(),
      created_at: a.created_at ? new Date(a.created_at) : new Date(0),
    }))
    .sort((a, b) => a.created_at.getTime() - b.created_at.getTime());

  // Explicit terminal decisions on PR by purchasing.
  const purchasingFinal = normalizedApprovals.filter(
    (a) => a.role === "purchasing_final" && (a.status === "rejected" || a.status === "cancelled")
  );
  if (purchasingFinal.length > 0) {
    return purchasingFinal[purchasingFinal.length - 1].status;
  }

  // Explicit requester cancellation is terminal.
  const requesterCancelled = normalizedApprovals.some(
    (a) => a.role === "requester_cancel" && a.status === "cancelled"
  );
  if (requesterCancelled) return "cancelled";

  // Any workflow rejection means rejected.
  const hasRejected = normalizedApprovals.some((a) => a.status === "rejected");
  if (hasRejected) return "rejected";

  const stepsLen = parseStepsLen(req.request_steps, req.workflow_steps);
  const approvedStepSet = new Set<number>();
  for (const a of normalizedApprovals) {
    if (a.status !== "approved") continue;
    if (a.step_index < 0) continue;
    approvedStepSet.add(a.step_index);
  }

  // Conversion only happens from approved PR. If linked PO exists and no terminal decision above, PR is approved.
  if (isPRName(name) && req.converted_po_request_id != null && Number(req.converted_po_request_id) > 0) {
    return "approved";
  }

  if (stepsLen > 0) {
    let allApproved = true;
    for (let i = 0; i < stepsLen; i += 1) {
      if (!approvedStepSet.has(i)) {
        allApproved = false;
        break;
      }
    }
    return allApproved ? "approved" : "pending";
  }

  // Fallback: keep current if we cannot infer.
  const cur = String(req.status || "").trim().toLowerCase();
  return cur || "pending";
}

async function main() {
  if (!conn) {
    console.error("Set SQLSERVER_CONNECTION_STRING in .env");
    process.exit(1);
  }

  const pool = await sql.connect(conn);
  console.log("[backfill] Connected to SQL Server");

  await pool.request().query(`
    IF COL_LENGTH('workflow_requests', 'po_status') IS NULL
      ALTER TABLE workflow_requests ADD po_status NVARCHAR(50) NULL;
  `);

  const reqRs = await pool.request().query(`
    SELECT
      r.id,
      r.status,
      r.template_name,
      w.category AS template_category,
      r.request_steps,
      w.steps AS workflow_steps,
      r.converted_po_request_id
    FROM workflow_requests r
    INNER JOIN workflows w ON w.id = r.template_id
    WHERE w.category = N'procurement';
  `);
  const requests: ReqRow[] = reqRs.recordset || [];
  if (requests.length === 0) {
    console.log("[backfill] No procurement requests found.");
    await pool.close();
    return;
  }

  const apRs = await pool.request().query(`
    SELECT request_id, step_index, status, approver_role_snapshot, created_at
    FROM request_approvals;
  `);
  const allApprovals: ApprovalRow[] = apRs.recordset || [];
  const approvalMap = new Map<number, ApprovalRow[]>();
  for (const ap of allApprovals) {
    const arr = approvalMap.get(Number(ap.request_id)) || [];
    arr.push(ap);
    approvalMap.set(Number(ap.request_id), arr);
  }

  let statusUpdated = 0;
  for (const req of requests) {
    const expected = computeExpectedStatus(req, approvalMap.get(Number(req.id)) || []);
    if (!expected) continue;
    const current = String(req.status || "").trim().toLowerCase();
    if (expected !== current) {
      await pool
        .request()
        .input("id", sql.Int, Number(req.id))
        .input("status", sql.NVarChar, expected)
        .query("UPDATE workflow_requests SET status = @status WHERE id = @id");
      statusUpdated += 1;
    }
  }

  const poSync = await pool.request().query(`
    UPDATE pr
    SET pr.po_status = po.status
    FROM workflow_requests pr
    INNER JOIN workflow_requests po ON po.id = pr.converted_po_request_id
    WHERE pr.converted_po_request_id IS NOT NULL
      AND (
        pr.po_status IS NULL
        OR LOWER(LTRIM(RTRIM(CAST(pr.po_status AS NVARCHAR(50))))) <> LOWER(LTRIM(RTRIM(CAST(po.status AS NVARCHAR(50)))))
      );
    SELECT @@ROWCOUNT AS updated_count;
  `);
  const poStatusUpdated = Number(poSync.recordset?.[0]?.updated_count ?? 0);

  console.log(`[backfill] workflow_requests.status updated: ${statusUpdated}`);
  console.log(`[backfill] workflow_requests.po_status synced: ${poStatusUpdated}`);

  await pool.close();
  console.log("[backfill] Done.");
}

main().catch((e) => {
  console.error("[backfill] Failed:", e);
  process.exit(1);
});

