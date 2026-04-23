-- Cost / project centers (MySQL). Run against your approval-system database.
-- This app’s Node server uses SQL Server by default; use this script only if you keep a parallel MySQL catalog.

CREATE TABLE IF NOT EXISTS cost_centers (
  id INT AUTO_INCREMENT PRIMARY KEY,
  code VARCHAR(64) NOT NULL,
  name VARCHAR(500) NOT NULL,
  gl_account VARCHAR(128) NULL,
  approval VARCHAR(512) NULL,
  status TINYINT(1) NOT NULL DEFAULT 1 COMMENT '1 = active, 0 = inactive',
  created_at DATETIME(3) NOT NULL DEFAULT (CURRENT_TIMESTAMP(3)),
  UNIQUE KEY uq_cost_centers_code (code),
  KEY ix_cost_centers_status (status)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- Optional seed (run once; skip rows that already exist)
-- INSERT INTO cost_centers (code, name, status) VALUES ('1000', 'IT Department', 1);
