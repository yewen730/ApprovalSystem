export interface Role {
  id: number;
  name: string;
  permissions: string[]; // JSON array of permission keys
  created_at: string;
}

export interface User {
  id: number;
  username: string;
  roles: string[];
  permissions: string[];
  department: string;
  /** Job title / designation from `usersetting.designation` (optional). */
  designation?: string | null;
  entities: string[]; // List of entities the user belongs to
  /** Max request total (MYR equivalent) this user may approve as "approver"; null/undefined = no per-user cap */
  approval_limit_myr?: number | null;
}

export interface WorkflowStep {
  id: string;
  label: string;
  approverRole: string;
}

export interface Workflow {
  id: number;
  creator_id: number;
  creator_name?: string;
  name: string;
  category: 'Purchasing' | 'HR' | 'Finance' | 'IT' | 'General'|'Shipping';
  steps: WorkflowStep[];
  table_columns?: string[];
  attachments_required?: boolean;
  is_active?: boolean;
  status: 'Pending' | 'Approved' | 'Rejected';
  created_at: string;
}

export interface Attachment {
  id: number;
  file_name: string;
  file_type: string;
  /** Legacy inline base64 / data URL when row was stored in DB. */
  file_data?: string | null;
  /** Relative path under company attachment storage (server only). */
  file_path?: string | null;
  /** Authenticated download/view URL (preferred when files are on disk). */
  file_url?: string | null;
}

export interface WorkflowRequest {
  id: number;
  template_id: number;
  requester_id: number;
  requester_name: string;
  department: string;
  template_name: string;
  template_steps: WorkflowStep[];
  table_columns?: string[];
  category?: string;
  attachments_required?: boolean;
  title: string;
  details: string;
  entity?: string; // The entity this request belongs to
  /** GCCM: requester-selected user id for approver-role steps (see server rules). */
  assigned_approver_id?: number | null;
  /** List API: request total (MYR) when pending on approver step (for limit / escalation UI). */
  request_total_myr_snapshot?: number | null;
  /** Persisted request total amount in MYR (stored in SQL workflow_requests.total_amount_myr). */
  total_amount_myr?: number | null;
  /** List API: chosen approver's cap (MYR) for this entity; null = no cap in DB. */
  assigned_approver_limit_myr?: number | null;
  /** List API: true when amount exceeds chosen approver's cap (another same-dept higher-limit approver may sign). */
  assigned_approval_shortfall?: boolean;
  formatted_id?: string; // The auto-generated PR ID
  currency?: string;
  cost_center?: string;
  section?: string;
  /** One supplier for the whole PR (not per line). */
  suggested_supplier?: string | null;
  /** Set when this PR has been converted to a PO (server); one PR → one PO. */
  converted_po_request_id?: number | null;
  /** Populated on list API for PR rows: linked PO document id (official PO number). */
  linked_po_formatted_id?: string | null;
  /** Populated on list API for PR rows: workflow status of the linked PO request. */
  linked_po_status?: string | null;
  line_items?: any[];
  tax_rate?: number;
  /** Document-level discount 0–1 (e.g. 0.05 = 5% off subtotal before tax). PR-focused; PO may carry from PR convert. */
  discount_rate?: number | null;
  status: 'Pending' | 'Approved' | 'Rejected' | 'Cancelled';
  current_step_index: number;
  /** Requester pad image as data URL; persisted in SQL as NVARCHAR(MAX) for PDFs. */
  requester_signature?: string;
  /** Timestamp when the requester completed the signature pad (image also in requester_signature when stored). */
  requester_signed_at?: string | null;
  /** From `usersetting.designation` join (PDF / display). */
  requester_designation?: string | null;
  created_at: string;
  approvals?: RequestApproval[];
}

export interface RequestApproval {
  id: number;
  request_id: number;
  step_index: number;
  approver_id: number;
  approver_name: string;
  /** Workflow step role when this approval was recorded (e.g. checker, approver). */
  approver_role_snapshot?: string | null;
  /** From `usersetting.designation` at approval time (PDF / display). */
  approver_designation?: string | null;
  status: 'Approved' | 'Rejected' | 'Cancelled';
  comment: string;
  /** Approver pad image as data URL; NVARCHAR(MAX) for PDFs. */
  approver_signature?: string;
  /** When set, a delegate (e.g. sign-on-behalf) recorded the signature for `approver_id`. */
  signed_by_user_id?: number | null;
  /** Display name of the user who signed on behalf (from server join). */
  signed_by_name?: string | null;
  created_at: string;
}
