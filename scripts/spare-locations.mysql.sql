-- Entity-scoped spare location master list for SR "Spare for (Location)" dropdown.
-- Note: the running app currently reads SQL Server endpoints; keep this MySQL table
-- in sync only if your organization maintains a separate MySQL master catalog.

CREATE TABLE IF NOT EXISTS spare_locations (
  id INT NOT NULL AUTO_INCREMENT,
  entity VARCHAR(32) NOT NULL,
  code VARCHAR(64) NOT NULL,
  name VARCHAR(500) NOT NULL,
  status TINYINT(1) NOT NULL DEFAULT 1,
  created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (id),
  UNIQUE KEY uq_spare_locations_entity_code (entity, code),
  KEY ix_spare_locations_entity_status (entity, status, code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Optional example seed rows (edit for your entities).
-- INSERT INTO spare_locations (entity, code, name, status) VALUES
-- ('GCCM', 'LOC-RAW', 'Raw Material Store', 1),
-- ('GCCM', 'LOC-MNT', 'Maintenance Store', 1),
-- ('GCIB', 'LOC-MAIN', 'Main Warehouse', 1);
