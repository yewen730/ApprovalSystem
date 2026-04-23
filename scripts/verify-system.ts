/**
 * Step-by-step checks: .env → SQL Server → core tables used by the app.
 * Run from project root: npm run verify:system
 */
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";
import sql from "mssql";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.join(__dirname, "..", ".env") });

async function main() {
  console.log("=== Flowmaster system verification ===\n");

  console.log("Step 1: Load .env from project root");
  const cs = process.env.SQLSERVER_CONNECTION_STRING?.trim();
  if (!cs) {
    console.error("  FAIL: SQLSERVER_CONNECTION_STRING is missing or empty in .env");
    process.exit(1);
  }
  console.log("  OK: Connection string is set (credentials not printed).\n");

  console.log("Step 2: Connect to SQL Server");
  let pool: sql.ConnectionPool;
  try {
    pool = await sql.connect(cs);
    console.log("  OK: Connected.\n");
  } catch (e: any) {
    console.error("  FAIL:", e?.message || e);
    process.exit(1);
  }

  const coreTables = [
    "workflow_requests",
    "workflows",
    "users",
    "request_attachments",
    "request_approvals",
    "attachments",
  ];

  console.log("Step 3: Core application tables");
  for (const t of coreTables) {
    try {
      const rs = await pool.request().query(`
        SELECT OBJECT_ID(N'dbo.${t}', N'U') AS table_id
      `);
      const id = rs.recordset[0]?.table_id;
      if (id == null) {
        console.log(`  WARN: dbo.${t} not found`);
      } else {
        console.log(`  OK: dbo.${t} (object_id: ${id})`);
      }
    } catch (e: any) {
      console.error(`  FAIL (${t}):`, e?.message || e);
    }
  }
  console.log("");

  console.log("Step 4: User count (sanity check)");
  try {
    const rs = await pool.request().query(`SELECT COUNT(*) AS cnt FROM dbo.users`);
    console.log("  users row count:", rs.recordset[0]?.cnt ?? "?");
  } catch (e: any) {
    console.log("  SKIP (users table may differ):", e?.message || e);
  }

  await pool.close();
  console.log("\n=== Done ===");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
