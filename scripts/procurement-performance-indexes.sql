-- Optional SQL Server indexes to speed up Procurement Center list and approval lookups.
-- Review with your DBA; run during a maintenance window on busy databases.
-- If an index already exists with the same definition, skip that block.

-- Procurement list: filter by entity (or IN list) and sort by created_at.
IF NOT EXISTS (
  SELECT 1 FROM sys.indexes i
  INNER JOIN sys.tables t ON i.object_id = t.object_id
  INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
  WHERE s.name = N'dbo' AND t.name = N'workflow_requests' AND i.name = N'IX_workflow_requests_entity_created_at'
)
BEGIN
  CREATE NONCLUSTERED INDEX IX_workflow_requests_entity_created_at
    ON dbo.workflow_requests (entity, created_at DESC)
    INCLUDE (template_id, status, department, requester_id, assigned_approver_id);
END
GO

-- Join workflow_requests to workflows by template; optional if template_id is already well-indexed.
IF NOT EXISTS (
  SELECT 1 FROM sys.indexes i
  INNER JOIN sys.tables t ON i.object_id = t.object_id
  INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
  WHERE s.name = N'dbo' AND t.name = N'workflows' AND i.name = N'IX_workflows_category_id'
)
BEGIN
  CREATE NONCLUSTERED INDEX IX_workflows_category_id
    ON dbo.workflows (category, id);
END
GO

-- Per-request approval rows (detail endpoint and PDFs).
IF NOT EXISTS (
  SELECT 1 FROM sys.indexes i
  INNER JOIN sys.tables t ON i.object_id = t.object_id
  INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
  WHERE s.name = N'dbo' AND t.name = N'request_approvals' AND i.name = N'IX_request_approvals_request_id_created_at'
)
BEGIN
  CREATE NONCLUSTERED INDEX IX_request_approvals_request_id_created_at
    ON dbo.request_approvals (request_id, created_at);
END
GO
