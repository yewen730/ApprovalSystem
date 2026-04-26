import "dotenv/config";
import sql from "mssql";

const cs = process.env.SQLSERVER_CONNECTION_STRING;
if (!cs) {
  console.error("Missing SQLSERVER_CONNECTION_STRING");
  process.exit(1);
}

const pool = await sql.connect(cs);
try {
  const counts = await pool.request().query(`
    SELECT
      UPPER(LTRIM(RTRIM(entity))) AS entity,
      COUNT(*) AS total,
      SUM(CASE WHEN status = 1 THEN 1 ELSE 0 END) AS active
    FROM dbo.cost_center
    GROUP BY UPPER(LTRIM(RTRIM(entity)))
    ORDER BY entity ASC;
  `);
  console.log("Counts by entity:", counts.recordset);

  const dupes = await pool.request().query(`
    SELECT TOP 20
      UPPER(LTRIM(RTRIM(entity))) AS entity,
      UPPER(LTRIM(RTRIM(code))) AS code,
      COUNT(*) AS c
    FROM dbo.cost_center
    GROUP BY UPPER(LTRIM(RTRIM(entity))), UPPER(LTRIM(RTRIM(code)))
    HAVING COUNT(*) > 1
    ORDER BY c DESC, entity ASC, code ASC;
  `);
  console.log("Top duplicates:", dupes.recordset);

  const sample = await pool.request().query(`
    SELECT TOP 20 entity, code, name, gl_account, status, created_at
    FROM dbo.cost_center
    WHERE status = 1
    ORDER BY entity ASC, code ASC;
  `);
  console.log("Sample active rows:", sample.recordset);
} finally {
  await pool.close();
}

