-- Per-entity approver registry: caps (GCBCM, CCI, ACI, …) + GCCM requestor picker list.
-- Username lives on [dbo].[usersetting]; this table stores user_id + entity + flags.

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
GO

-- Example: GCCM — show in requestor dropdown (no MYR cap enforced in app for GCCM).
/*
MERGE dbo.entity_approver_registry AS t
USING (SELECT id AS user_id FROM dbo.usersetting WHERE username = N'jdoe') AS s
ON t.entity = N'GCCM' AND t.user_id = s.user_id
WHEN MATCHED THEN UPDATE SET selectable_by_requestor = 1, active = 1, approval_limit_myr = NULL
WHEN NOT MATCHED THEN INSERT (entity, user_id, selectable_by_requestor, approval_limit_myr, active)
VALUES (N'GCCM', s.user_id, 1, NULL, 1);
*/

-- Example: GCBCM — cap only (not in GCCM-style picker unless you also set selectable_by_requestor for that entity).
/*
MERGE dbo.entity_approver_registry AS t
USING (SELECT id AS user_id FROM dbo.usersetting WHERE username = N'jdoe') AS s
ON t.entity = N'GCBCM' AND t.user_id = s.user_id
WHEN MATCHED THEN UPDATE SET approval_limit_myr = 50000.00, active = 1, selectable_by_requestor = 0
WHEN NOT MATCHED THEN INSERT (entity, user_id, selectable_by_requestor, approval_limit_myr, active)
VALUES (N'GCBCM', s.user_id, 0, 50000.00, 1);
*/
