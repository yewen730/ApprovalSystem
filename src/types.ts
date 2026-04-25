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
  formatted_id?: string; // The auto-generated PR ID
  currency?: string;
  cost_center?: string;
  section?: string;
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
  created_at: string;
  approvals?: RequestApproval[];
}

export interface RequestApproval {
  id: number;
  request_id: number;
  step_index: number;
  approver_id: number;
  approver_name: string;
  status: 'Approved' | 'Rejected' | 'Cancelled';
  comment: string;
  /** Approver pad image as data URL; NVARCHAR(MAX) for PDFs. */
  approver_signature?: string;
  created_at: string;
}
