import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { buildProcurementPrFormPdfDoc } from './procurement/prFormPdfGenerator';
import { 
  LayoutDashboard, 
  PlusCircle, 
  Users, 
  CheckCircle2, 
  XCircle, 
  FileText, 
  Upload, 
  LogOut, 
  Shield, 
  ChevronRight,
  Paperclip,
  Trash2,
  Clock,
  AlertCircle,
  Send,
  Inbox,
  ClipboardList,
  Filter,
  Building2,
  Download,
  RotateCcw,
  Edit2,
  UserPlus,
  ShoppingCart,
  Package,
  Warehouse,
  Receipt,
  Search,
  ChevronDown,
  Plus,
  X,
  Eye,
  RefreshCw,
  TrendingUp,
  Activity,
  FileSpreadsheet,
  ChevronsLeft,
  ChevronsRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { Toaster, toast } from 'react-hot-toast';
import { jsPDF } from 'jspdf';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell
} from 'recharts';
import { User, Workflow, WorkflowStep, Attachment, WorkflowRequest, RequestApproval, Role } from './types';
import { cn } from './lib/utils';

// --- API Helpers ---
const FLOWMASTER_ENTITY_KEY = 'flowmaster_entity';
const FLOWMASTER_SIDEBAR_COLLAPSED_KEY = 'flowmaster_sidebar_collapsed';
const FLOWMASTER_ENTITY_CHANGED_EVENT = 'flowmaster-entity-changed';

const api = {
  token: localStorage.getItem('token'),
  setToken(token: string | null) {
    this.token = token;
    if (token) localStorage.setItem('token', token);
    else {
      localStorage.removeItem('token');
      localStorage.removeItem(FLOWMASTER_ENTITY_KEY);
    }
  },
  setActiveEntity(entity: string | null) {
    if (entity) localStorage.setItem(FLOWMASTER_ENTITY_KEY, entity);
    else localStorage.removeItem(FLOWMASTER_ENTITY_KEY);
    window.dispatchEvent(new CustomEvent(FLOWMASTER_ENTITY_CHANGED_EVENT, { detail: { entity } }));
  },
  async request(path: string, options: RequestInit & { skipEntity?: boolean } = {}) {
    const { skipEntity: skipEntityOpt, ...fetchOptions } = options;
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${this.token}`,
      ...(fetchOptions.headers as Record<string, string>),
    };
    const skipEntity = !!skipEntityOpt || path === '/api/login' || path === '/api/register' || path === '/api/me';
    const entity = localStorage.getItem(FLOWMASTER_ENTITY_KEY);
    if (!skipEntity && entity) {
      headers['X-Entity'] = entity;
    }
    const res = await fetch(path, {
      ...fetchOptions,
      headers,
    });
    const contentType = res.headers.get('content-type') || '';
    const isJson = contentType.includes('application/json');

    if (!res.ok) {
      if (isJson) {
        const err = await res.json();
        const base = err.error || 'Something went wrong';
        const detail = err.details != null && String(err.details).trim() ? ` — ${String(err.details).trim()}` : '';
        throw new Error(base + detail);
      }
      const text = await res.text();
      throw new Error(text || `Request failed (${res.status})`);
    }

    if (isJson) return res.json();
    // Some endpoints might return empty responses (e.g., 204) or non-JSON.
    // Keep callers safe by returning a minimal object.
    const text = await res.text().catch(() => '');
    return text ? { message: text } : {};
  }
};

const MALAYSIA_TIME_ZONE = 'Asia/Kuala_Lumpur';

function formatDateMYT(value: string | number | Date | null | undefined): string {
  if (value == null || value === '') return '';
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleDateString(undefined, { timeZone: MALAYSIA_TIME_ZONE });
}

function formatDateTimeMYT(value: string | number | Date | null | undefined): string {
  if (value == null || value === '') return '';
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleString(undefined, { timeZone: MALAYSIA_TIME_ZONE });
}

/** Open or download a request attachment (disk-backed `file_url` or legacy `file_data`). */
async function openWorkflowRequestAttachment(
  requestId: number,
  att: Attachment,
  setViewingPdf: (pdf: { url: string; fileName: string } | null) => void
): Promise<void> {
  const isPdf = att.file_type === "application/pdf" || att.file_name.toLowerCase().endsWith(".pdf");
  const headers: Record<string, string> = { Authorization: `Bearer ${api.token}` };
  const entity = localStorage.getItem(FLOWMASTER_ENTITY_KEY);
  if (entity) headers["X-Entity"] = entity;
  let url: string | undefined;
  if (att.file_data && String(att.file_data).trim()) {
    url = att.file_data;
  } else {
    const fetchPath =
      att.file_url && att.file_url.startsWith("/")
        ? att.file_url
        : att.id
          ? `/api/workflow-requests/${requestId}/attachments/${att.id}/file`
          : null;
    if (!fetchPath) {
      toast.error("File not available");
      return;
    }
    const res = await fetch(fetchPath, { headers });
    if (!res.ok) {
      toast.error("Could not open file");
      return;
    }
    const blob = await res.blob();
    url = URL.createObjectURL(blob);
  }
  if (isPdf) setViewingPdf({ url, fileName: att.file_name });
  else {
    const link = document.createElement("a");
    link.href = url;
    link.download = att.file_name;
    link.click();
    if (url.startsWith("blob:")) URL.revokeObjectURL(url);
  }
}

function downloadFileFromUrl(url: string, fileName: string): void {
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
}

// --- Components ---

const PdfViewer = ({
  url,
  fileName,
  onDownload,
  onClose,
}: {
  url: string;
  fileName: string;
  onDownload: () => void;
  onClose: () => void;
}) => {
  const iframeSrc = `${url}${url.includes("#") ? "&" : "#"}view=FitH`;
  return (
    <div className="fixed inset-0 z-[100] bg-black/60 backdrop-blur-sm p-2 md:p-3">
      <motion.div
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className="relative w-full h-full bg-white rounded-2xl shadow-2xl overflow-hidden flex flex-col"
      >
        <div className="flex items-center justify-between gap-3 p-4 border-b border-zinc-100 bg-white">
          <h3 className="text-sm font-bold text-zinc-900 truncate">PDF Viewer: {fileName}</h3>
          <div className="flex items-center gap-2">
            <button
              onClick={onDownload}
              className="inline-flex items-center gap-2 px-3 py-2 text-xs font-bold rounded-lg bg-indigo-600 text-white hover:bg-indigo-500 transition-colors"
            >
              <Download className="w-4 h-4" />
              Download
            </button>
            <button
              onClick={onClose}
              className="p-2 hover:bg-zinc-100 rounded-full transition-colors"
            >
              <X className="w-5 h-5 text-zinc-500" />
            </button>
          </div>
        </div>
        <div className="flex-1 bg-zinc-100 overflow-hidden">
          <iframe
            src={iframeSrc}
            className="w-full h-full border-none"
            title="PDF Viewer"
          />
        </div>
      </motion.div>
    </div>
  );
};

function buildPdfPreview(doc: jsPDF, fileName: string): { url: string; fileName: string; pdfDataUrl: string } {
  const pdfDataUrl = doc.output("dataurlstring");
  const blob = doc.output("blob");
  const url = URL.createObjectURL(blob);
  return { url, fileName, pdfDataUrl };
}

/** Saves PR/SR/PO form PDF next to quotation uploads (same UNC folder per request). */
async function persistGeneratedProcurementFormPdf(requestId: number, pdfDataUrl: string): Promise<void> {
  if (!requestId || !pdfDataUrl) return;
  try {
    await api.request(`/api/workflow-requests/${requestId}/generated-form-pdf`, {
      method: "POST",
      body: JSON.stringify({ pdf_data: pdfDataUrl }),
    });
  } catch (e) {
    console.error("Failed to archive form PDF to server:", e);
  }
}

function toUpperSerial(value: string | number | null | undefined): string {
  return String(value ?? "").toUpperCase();
}

function displayRequestSerial(request: Pick<WorkflowRequest, "formatted_id" | "id">): string {
  return toUpperSerial(request.formatted_id || request.id);
}

const SECTION_OPTIONS = ["NAR1", "NAR2", "POWDER PLANT", "PRESS", "INSTRUMENT", "FACILITY", "NA"] as const;
const SECTION_NA = "NA";

function sectionSelectionFromStored(value: string | null | undefined): string {
  const normalized = String(value ?? "").trim().toUpperCase();
  return SECTION_OPTIONS.includes(normalized as (typeof SECTION_OPTIONS)[number]) ? normalized : "";
}

function sectionPayloadFromSelection(value: string | null | undefined): string {
  const normalized = String(value ?? "").trim().toUpperCase();
  return normalized === SECTION_NA ? "" : normalized;
}

type ConvertPoModalTarget = {
  id: number;
  title: string;
  prSerial: string;
  entityCode: string;
};

type ConvertPoUploadPayload = {
  name: string;
  type: string;
  data: string;
};

type PurchasingDecisionModalTarget = {
  id: number;
  decision: 'cancelled' | 'rejected';
  prSerial: string;
  title: string;
  entity: string;
};

const PurchasingDecisionModal = ({
  target,
  loading,
  onClose,
  onConfirm,
}: {
  target: PurchasingDecisionModalTarget | null;
  loading: boolean;
  onClose: () => void;
  onConfirm: (comment: string) => void;
}) => {
  const [comment, setComment] = useState('');
  useEffect(() => {
    if (target) setComment('');
  }, [target?.id, target?.decision]);

  if (!target) return null;
  const isCancel = target.decision === 'cancelled';
  return (
    <div
      className="fixed inset-0 z-[80] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm"
      role="dialog"
      aria-modal="true"
      aria-labelledby="purchasing-decision-title"
      onClick={loading ? undefined : onClose}
    >
      <motion.div
        initial={{ opacity: 0, scale: 0.96 }}
        animate={{ opacity: 1, scale: 1 }}
        className="bg-white rounded-2xl shadow-2xl border border-zinc-200 w-full max-w-lg overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <div className={cn("px-6 py-5 border-b", isCancel ? "bg-rose-50 border-rose-100" : "bg-rose-50 border-rose-100")}>
          <div className="flex items-start gap-3">
            <div className={cn("w-10 h-10 rounded-xl flex items-center justify-center shrink-0",
              isCancel ? "bg-rose-600 shadow-lg shadow-rose-200" : "bg-rose-600 shadow-lg shadow-rose-200"
            )}>
              <XCircle className="w-5 h-5 text-white" />
            </div>
            <div className="min-w-0">
              <h3 id="purchasing-decision-title" className="text-lg font-black text-zinc-900">
                {isCancel ? 'Cancel PR (Final)' : 'Reject PR'}
              </h3>
              <p className="text-sm text-zinc-600 mt-1">
                PR <span className="font-mono font-semibold text-zinc-900">{target.prSerial}</span>
                {target.title ? ` — ${target.title}` : ''}
              </p>
              <p className="text-xs text-zinc-500 mt-1">
                {isCancel
                  ? 'Cancellation is final (cannot be resubmitted). A reason is required for audit.'
                  : 'Rejection returns the PR to the requester. They may edit and resubmit.'}
              </p>
            </div>
          </div>
        </div>

        <div className="p-6 space-y-2">
          <label className="text-sm font-medium text-zinc-700">
            {isCancel ? 'Cancellation reason' : 'Rejection reason'}
          </label>
          <textarea
            rows={4}
            className={cn(
              "w-full px-4 py-3 rounded-xl border text-sm outline-none focus:ring-2 bg-white resize-none",
              isCancel ? "border-rose-200 focus:ring-rose-500" : "border-rose-200 focus:ring-rose-500"
            )}
            placeholder={isCancel ? 'Please describe why this PR is cancelled…' : 'Optional…'}
            value={comment}
            onChange={(e) => setComment(e.target.value)}
            disabled={loading}
          />
        </div>

        <div className="px-6 pb-6 flex justify-end gap-2">
          <button
            type="button"
            disabled={loading}
            onClick={onClose}
            className="px-4 py-2 rounded-xl text-sm font-semibold text-zinc-600 hover:bg-zinc-100"
          >
            Back
          </button>
          <button
            type="button"
            disabled={loading}
            onClick={() => {
              const t = comment.trim();
              if (isCancel && !t) {
                toast.error('Please enter a cancellation reason.');
                return;
              }
              onConfirm(t);
            }}
            className={cn(
              "px-4 py-2 rounded-xl text-sm font-bold text-white disabled:opacity-50",
              isCancel ? "bg-rose-600 hover:bg-rose-700" : "bg-rose-600 hover:bg-rose-700"
            )}
          >
            {loading ? 'Saving…' : isCancel ? 'Confirm cancellation' : 'Confirm rejection'}
          </button>
        </div>
      </motion.div>
    </div>
  );
};

/** Purchasing enters the official PO number; creates a new PO or appends this PR into an existing pending PO in the same entity. */
const ConvertPrToPoModal = ({
  target,
  loading,
  onClose,
  onConfirm,
}: {
  target: ConvertPoModalTarget | null;
  loading: boolean;
  onClose: () => void;
  onConfirm: (poNumber: string, upload?: ConvertPoUploadPayload) => void;
}) => {
  const [poNumber, setPoNumber] = useState('');
  const [poUpload, setPoUpload] = useState<ConvertPoUploadPayload | null>(null);
  useEffect(() => {
    if (target) {
      setPoNumber('');
      setPoUpload(null);
    }
  }, [target?.id]);

  if (!target) return null;

  return (
    <div
      className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm"
      role="dialog"
      aria-modal="true"
      aria-labelledby="convert-po-title"
      onClick={loading ? undefined : onClose}
    >
      <motion.div
        initial={{ opacity: 0, scale: 0.96 }}
        animate={{ opacity: 1, scale: 1 }}
        className="bg-white rounded-2xl shadow-2xl border border-zinc-200 w-full max-w-md p-6"
        onClick={(e) => e.stopPropagation()}
      >
        <h3 id="convert-po-title" className="text-lg font-bold text-zinc-900">
          Record PO number
        </h3>
        <p className="text-sm text-zinc-600 mt-2">
          Converting PR <span className="font-mono font-semibold text-zinc-900">{target.prSerial}</span>
          {target.title ? ` — ${target.title}` : ''}
        </p>
        <p className="text-xs text-zinc-500 mt-1">
          Entity: <span className="font-semibold">{target.entityCode || '—'}</span>
        </p>
        <label htmlFor="official-po-number" className="block text-sm font-medium text-zinc-700 mt-5">
          Official PO number
        </label>
        <input
          id="official-po-number"
          type="text"
          autoComplete="off"
          className="mt-1.5 w-full px-4 py-2.5 rounded-xl border border-zinc-300 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
          placeholder="e.g. POGC26-00042"
          value={poNumber}
          onChange={(e) => setPoNumber(e.target.value)}
          disabled={loading}
        />
        <p className="text-xs text-zinc-500 mt-2">
          Required. If this PO number already exists as a pending PO in the same entity, this PR will be appended into that PO.
        </p>
        <label htmlFor="official-po-upload" className="block text-sm font-medium text-zinc-700 mt-5">
          PO document upload (optional)
        </label>
        <input
          id="official-po-upload"
          type="file"
          className="mt-1.5 block w-full text-sm text-zinc-700 file:mr-3 file:px-3 file:py-1.5 file:rounded-lg file:border-0 file:bg-zinc-100 file:text-zinc-700 hover:file:bg-zinc-200"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (!file) {
              setPoUpload(null);
              return;
            }
            const reader = new FileReader();
            reader.onload = () => {
              const data = typeof reader.result === 'string' ? reader.result : '';
              if (!data) {
                toast.error('Failed to read selected file.');
                return;
              }
              setPoUpload({
                name: file.name,
                type: file.type || 'application/octet-stream',
                data,
              });
            };
            reader.onerror = () => {
              toast.error('Failed to read selected file.');
              setPoUpload(null);
            };
            reader.readAsDataURL(file);
          }}
          disabled={loading}
        />
        <p className="text-xs text-zinc-500 mt-2">
          Optional. You can upload now or add the PO file later from the PO request details.
        </p>
        {poUpload ? (
          <p className="text-xs text-zinc-600 mt-1">
            Selected: <span className="font-medium">{poUpload.name}</span>
          </p>
        ) : null}
        <div className="flex justify-end gap-2 mt-6">
          <button
            type="button"
            disabled={loading}
            onClick={onClose}
            className="px-4 py-2 rounded-xl text-sm font-semibold text-zinc-600 hover:bg-zinc-100"
          >
            Cancel
          </button>
          <button
            type="button"
            disabled={loading}
            onClick={() => {
              const t = poNumber.trim();
              if (!t) {
                toast.error('Enter the official PO number.');
                return;
              }
              onConfirm(t, poUpload || undefined);
            }}
            className="px-4 py-2 rounded-xl text-sm font-bold text-white bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50"
          >
            {loading ? 'Processing…' : 'Create or append PO'}
          </button>
        </div>
      </motion.div>
    </div>
  );
};

const EntitySelection = ({ entities, onSelect }: { entities: string[]; onSelect: (entity: string) => void }) => {
  return (
    <div className="min-h-screen flex items-center justify-center bg-zinc-50 p-4">
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="w-full max-w-md bg-white rounded-2xl shadow-xl border border-zinc-200 p-8"
      >
        <div className="flex flex-col items-center mb-8">
          <div className="w-12 h-12 bg-indigo-600 rounded-xl flex items-center justify-center mb-4 shadow-lg shadow-indigo-200">
            <Building2 className="text-white w-6 h-6" />
          </div>
          <h1 className="text-2xl font-bold text-zinc-900">Select Entity</h1>
          <p className="text-zinc-500 text-sm mt-1">Please select an entity to continue</p>
        </div>

        <div className="grid gap-3">
          {entities.map((entity) => (
            <button
              key={entity}
              onClick={() => onSelect(entity)}
              className="w-full flex items-center justify-between p-4 rounded-xl border border-zinc-100 hover:border-indigo-200 hover:bg-indigo-50 transition-all group"
            >
              <span className="font-bold text-zinc-700 group-hover:text-indigo-600">{entity}</span>
              <ChevronRight className="w-5 h-5 text-zinc-300 group-hover:text-indigo-400" />
            </button>
          ))}
        </div>
      </motion.div>
    </div>
  );
};

const Login = ({ onLogin }: { onLogin: (user: User) => void }) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    try {
      const data = await api.request('/api/login', {
        method: 'POST',
        body: JSON.stringify({ username, password }),
      });
      api.setToken(data.token);
      onLogin(data.user);
      toast.success('Welcome back!');
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-zinc-50 p-4">
      <motion.div 
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="w-full max-w-md bg-white rounded-2xl shadow-xl border border-zinc-200 p-8"
      >
        <div className="flex flex-col items-center mb-8">
          <div className="w-12 h-12 bg-indigo-600 rounded-xl flex items-center justify-center mb-4 shadow-lg shadow-indigo-200">
            <Shield className="text-white w-6 h-6" />
          </div>
          <h1 className="text-2xl font-bold text-zinc-900">Approval System</h1>
          <p className="text-zinc-500 text-sm mt-1">
            Sign in to your account
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-zinc-700 mb-1">Username</label>
            <input
              type="text"
              required
              className="w-full px-4 py-2 rounded-lg border border-zinc-300 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-zinc-700 mb-1">Password</label>
            <input
              type="password"
              required
              className="w-full px-4 py-2 rounded-lg border border-zinc-300 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </div>
          <button
            type="submit"
            disabled={loading}
            className="w-full bg-indigo-600 text-white py-2 rounded-lg font-semibold hover:bg-indigo-700 transition-colors disabled:opacity-50"
          >
            {loading ? 'Processing...' : 'Login'}
          </button>
        </form>
      </motion.div>
    </div>
  );
};

const TemplateSelector = ({ templates, onSelect }: { templates: Workflow[], onSelect: (w: Workflow) => void }) => {
  // Backend stores workflow status in lowercase (e.g. "approved"), while UI types/UI code
  // sometimes expect Title Case. Normalize to make the submit page work reliably.
  const approvedTemplates = templates.filter(
    (t) =>
      (t.status ?? '').toString().trim().toLowerCase() === 'approved' &&
      (t.is_active ?? true)
  );

  return (
    <div className="space-y-6">
      <div className="flex flex-col gap-1">
        <h2 className="text-xl font-bold text-zinc-900">Select a Workflow Template</h2>
        <p className="text-sm text-zinc-500">Choose an approved template to start your request</p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {approvedTemplates.length === 0 && (
          <div className="col-span-full text-center py-12 bg-white rounded-xl border border-dashed border-zinc-300">
            <AlertCircle className="w-12 h-12 text-zinc-300 mx-auto mb-3" />
            <p className="text-zinc-500">No approved templates available. Please create one or wait for approval.</p>
          </div>
        )}
        {approvedTemplates.map((w) => (
          <div
            key={w.id}
            onClick={() => onSelect(w)}
            className="bg-white p-5 rounded-xl border border-zinc-200 hover:border-indigo-300 hover:shadow-md transition-all cursor-pointer group flex flex-col justify-between"
          >
            <div>
              <div className="flex items-center gap-2 mb-2">
                <div className="w-8 h-8 bg-indigo-50 rounded-lg flex items-center justify-center">
                  <ClipboardList className="w-4 h-4 text-indigo-600" />
                </div>
                <h3 className="font-bold text-zinc-900 group-hover:text-indigo-600 transition-colors">{w.name}</h3>
              </div>
            </div>
            <div className="flex items-center justify-between pt-4 border-t border-zinc-50">
              <span className="text-xs text-zinc-400 font-medium">{w.steps.length} Approval Steps</span>
              <div className="text-indigo-600 flex items-center gap-1 text-xs font-bold">
                Select <ChevronRight className="w-3 h-3" />
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

const WorkflowCreator = ({ onSuccess, availableRoles }: { onSuccess: () => void, availableRoles: Role[] }) => {
  const [name, setName] = useState('');
  const [category, setCategory] = useState<'Purchasing' | 'HR' | 'Finance' | 'IT' | 'General'>('general');
  const [steps, setSteps] = useState<WorkflowStep[]>([]);
  const [users, setUsers] = useState<User[]>([]);
  const [tableColumns, setTableColumns] = useState<{ id: string; name: string }[]>([]);
  const [newColumnName, setNewColumnName] = useState('');
  const [attachmentsRequired, setAttachmentsRequired] = useState(false);
  const [loading, setLoading] = useState(false);
  const isPR = isPRTemplate({ name, category });
  const isPO = isPOTemplate({ name, category });
  const isSR = isSRTemplate({ name, category });

  useEffect(() => {
    if (isPR) setSteps(FIXED_PR_STEPS);
    else if (isPO) setSteps(FIXED_PO_STEPS_FULL);
    else if (isSR) setSteps(FIXED_SR_STEPS);
  }, [isPR, isPO, isSR]);

  useEffect(() => {
    const fetchUsers = async () => {
      try {
        const data = await api.request('/api/users');
        setUsers(data);
      } catch (err) {
        console.error('Failed to fetch users', err);
      }
    };
    fetchUsers();
  }, []);

  useEffect(() => {
    if (steps.length === 0 && availableRoles.length > 0) {
      setSteps([{ 
        id: Math.random().toString(36).substr(2, 9), 
        label: 'Initial Review', 
        approverRole: availableRoles.find(r => r.name === 'checker')?.name || availableRoles[0].name 
      }]);
    }
  }, [availableRoles]);

  const addStep = () => {
    if (isPR || isPO || isSR) return;
    setSteps([...steps, { 
      id: Math.random().toString(36).substr(2, 9), 
      label: '', 
      approverRole: availableRoles.find(r => r.name === 'approver')?.name || availableRoles[0]?.name || '' 
    }]);
  };

  const removeStep = (id: string) => {
    if (isPR || isPO || isSR) return;
    setSteps(steps.filter(s => s.id !== id));
  };

  const addColumn = () => {
    if (newColumnName.trim()) {
      setTableColumns([...tableColumns, { id: Math.random().toString(36).substr(2, 9), name: newColumnName.trim() }]);
      setNewColumnName('');
    }
  };

  const removeColumn = (id: string) => {
    setTableColumns(tableColumns.filter(c => c.id !== id));
  };

  const updateColumn = (id: string, newName: string) => {
    setTableColumns(tableColumns.map(c => c.id === id ? { ...c, name: newName } : c));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (steps.some(s => !s.label || !s.approverRole)) {
      return toast.error('Please fill in all step details');
    }
    setLoading(true);
    try {
      await api.request('/api/workflows', {
        method: 'POST',
        body: JSON.stringify({ 
          name, 
          category,
          steps, 
          table_columns: tableColumns.map(c => c.name), 
          attachments_required: attachmentsRequired
        }),
      });
      toast.success('Workflow template submitted for approval!');
      onSuccess();
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="bg-white rounded-xl border border-zinc-200 p-6 shadow-sm">
      <h2 className="text-xl font-bold text-zinc-900 mb-6 flex items-center gap-2">
        <PlusCircle className="w-5 h-5 text-indigo-600" />
        Design New Workflow Template
      </h2>
      <form onSubmit={handleSubmit} className="space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="space-y-2">
            <label className="text-sm font-medium text-zinc-700">Workflow Name</label>
            <input
              type="text"
              required
              placeholder="e.g., Expense Reimbursement"
              className="w-full px-4 py-2 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
              value={name}
              onChange={(e) => setName(e.target.value)}
            />
          </div>
          <div className="space-y-2">
            <label className="text-sm font-medium text-zinc-700">Category</label>
            <select
              className="w-full px-4 py-2 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
              value={category}
              onChange={(e) => setCategory(e.target.value as any)}
            >
              <option value="general">General</option>
              <option value="procurement">Procurement</option>
              <option value="hr">HR</option>
              <option value="finance">Finance</option>
              <option value="it">IT</option>
            </select>
          </div>
        </div>

        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <label className="text-sm font-medium text-zinc-700">Approval Steps</label>
            {!isPR && !isPO && !isSR && (
              <button
                type="button"
                onClick={addStep}
                className="text-xs bg-indigo-50 text-indigo-600 px-3 py-1 rounded-full font-semibold hover:bg-indigo-100 transition-colors"
              >
                + Add Step
              </button>
            )}
          </div>
          <div className="space-y-3">
            {steps.map((step, index) => (
              <div key={step.id} className="flex items-center gap-3 bg-zinc-50 p-3 rounded-lg border border-zinc-200">
                <div className="w-6 h-6 bg-indigo-600 text-white rounded-full flex items-center justify-center text-xs font-bold">
                  {index + 1}
                </div>
                <input
                  placeholder="Step Label (e.g., Dept Head Approval)"
                  className="flex-1 bg-transparent border-none focus:ring-0 text-sm"
                  disabled={isPR || isPO || isSR}
                  value={step.label}
                  onChange={(e) => {
                    const newSteps = [...steps];
                    newSteps[index].label = e.target.value;
                    setSteps(newSteps);
                  }}
                />
                <select
                  className="w-32 bg-transparent border-none focus:ring-0 text-sm outline-none"
                  disabled={isPR || isPO || isSR}
                  value={step.approverRole}
                  onChange={(e) => {
                    const newSteps = [...steps];
                    newSteps[index].approverRole = e.target.value;
                    setSteps(newSteps);
                  }}
                >
                  {availableRoles.map(r => (
                    <option key={r.id} value={r.name}>{r.name.charAt(0).toUpperCase() + r.name.slice(1)}</option>
                  ))}
                </select>
                {steps.length > 1 && !isPR && !isPO && !isSR && (
                  <button type="button" onClick={() => removeStep(step.id)} className="text-zinc-400 hover:text-red-500">
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>

        <div className="flex items-center gap-2 bg-zinc-50 p-4 rounded-xl border border-zinc-200">
          <input
            type="checkbox"
            id="attachments_required"
            className="w-4 h-4 text-indigo-600 rounded border-zinc-300 focus:ring-indigo-500"
            checked={attachmentsRequired}
            onChange={(e) => setAttachmentsRequired(e.target.checked)}
          />
          <label htmlFor="attachments_required" className="text-sm font-medium text-zinc-700 cursor-pointer">
            Require attachments when submitting requests
          </label>
        </div>

        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <label className="text-sm font-medium text-zinc-700">Data Table Columns (Optional)</label>
            <div className="flex gap-2">
              <input
                type="text"
                placeholder="e.g., Amount, Qty"
                className="text-xs px-2 py-1 rounded border border-zinc-300 outline-none focus:ring-1 focus:ring-indigo-500"
                value={newColumnName}
                onChange={(e) => setNewColumnName(e.target.value)}
                onKeyPress={(e) => e.key === 'Enter' && (e.preventDefault(), addColumn())}
              />
              <button
                type="button"
                onClick={addColumn}
                className="text-xs bg-emerald-50 text-emerald-600 px-3 py-1 rounded-full font-semibold hover:bg-emerald-100 transition-colors"
              >
                + Add Column
              </button>
            </div>
          </div>
          <div className="flex flex-wrap gap-2">
            {tableColumns.length === 0 && <p className="text-xs text-zinc-400 italic">No custom columns defined. Requests will only have title and details.</p>}
            {tableColumns.map((col) => (
              <div key={col.id} className="flex items-center gap-2 bg-zinc-100 px-3 py-1 rounded-full text-xs text-zinc-600 border border-zinc-200 focus-within:border-indigo-300 focus-within:ring-1 focus-within:ring-indigo-500 transition-all">
                <input
                  type="text"
                  className="bg-transparent border-none focus:ring-0 text-xs w-24 outline-none font-medium"
                  value={col.name}
                  onChange={(e) => updateColumn(col.id, e.target.value)}
                  placeholder="Column name..."
                />
                <button type="button" onClick={() => removeColumn(col.id)} className="text-zinc-400 hover:text-red-500">
                  <Trash2 className="w-3 h-3" />
                </button>
              </div>
            ))}
          </div>
        </div>

        <button
          type="submit"
          disabled={loading}
          className="w-full bg-indigo-600 text-white py-3 rounded-xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 disabled:opacity-50"
        >
          {loading ? 'Submitting...' : 'Submit Workflow Template for Approval'}
        </button>
      </form>
    </div>
  );
};

/** Line-item column for short notes (legacy JSON may use "Remarks (Purpose)"). */
const REMARKS_LINE_COL = 'Remarks';

/** Entity code on requests / X-Entity → legal name on PR/PO PDFs and drafts. */
const ENTITY_LEGAL_NAMES: Record<string, string> = {
  GCBCM: 'GCB COCOA MALAYSIA SDN BHD',
  GCCM: 'GUAN CHONG COCOA MANUFACTURER SDN BHD',
  CCI: "GCB COCOA COTE D'IVOIRE",
  GCBCS: 'GCB COCOA SINGAPORE PTE LTD',
  ACI: 'PT. ACI COCOA INDONESIA',
};

function entityLegalDisplayName(code: string | undefined | null): string {
  if (code == null || !String(code).trim()) return '-';
  const key = String(code).trim().toUpperCase();
  return ENTITY_LEGAL_NAMES[key] || String(code).trim();
}

/** SST (Malaysia entities), TVA (CCI), generic Tax elsewhere — for PDFs and forms. */
function procurementTaxLabelForEntity(code: string | undefined | null): string {
  const k = String(code ?? '').trim().toUpperCase();
  if (k === 'GCCM' || k === 'GCBCM') return 'SST';
  if (k === 'CCI') return 'TVA';
  return 'Tax';
}

function procurementTaxRateFormLabel(code: string | undefined | null): string {
  const lab = procurementTaxLabelForEntity(code);
  return lab === 'Tax' ? 'Tax rate' : `${lab} rate`;
}

function clampUnitRate(n: number): number {
  if (Number.isNaN(n) || n < 0) return 0;
  if (n > 1) return 1;
  return n;
}

/** Round to 2 decimal places (currency sen). */
function roundMoney2(n: number): number {
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 100) / 100;
}

/** Locale grouping with exactly 2 fraction digits (avoids 3dp from raw floats / default toLocaleString). */
function formatProcurementMoney(n: number): string {
  return roundMoney2(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/** User enters percent (e.g. 10 = 10%); API/storage use unit rate 0–1. */
function procurementPercentToUnitRate(percent: number): number {
  return clampUnitRate(percent / 100);
}

/** Stored unit rate (0–1) → percent for form fields (e.g. 0.18 → 18). `whenMissingUseUnitRate` is the stored default when value is null/undefined. */
function procurementUnitRateToPercent(unit: number | undefined | null, whenMissingUseUnitRate: number): number {
  const u =
    unit !== undefined && unit !== null && !Number.isNaN(Number(unit))
      ? Number(unit)
      : whenMissingUseUnitRate;
  return Math.round(u * 10000) / 100;
}

/** Line qty for money: max quantity (stock requisition), else quantity, else min quantity — matches server `lineItemQtyForTotals`. */
function lineItemQtyForDisplay(item: any): number {
  const max = parseFloat(String(item?.['Max quantity'] ?? item?.['Max Quantity'] ?? ''));
  if (!Number.isNaN(max) && max > 0) return max;
  const q = parseFloat(String(item?.['Quantity'] ?? item?.quantity ?? 0));
  if (!Number.isNaN(q) && q > 0) return q;
  const min = parseFloat(String(item?.['Min quantity'] ?? item?.['Min Quantity'] ?? ''));
  if (!Number.isNaN(min)) return min;
  return 0;
}

/** Subtotal, discount (on subtotal), tax (on amount after discount), total — same order as server `computeRequestTotalMyr`; amounts rounded to 2 dp (sen). */
function procurementMoneyTotals(req: Pick<WorkflowRequest, 'line_items' | 'tax_rate' | 'discount_rate'>) {
  let subtotal = 0;
  for (const item of req.line_items || []) {
    const qty = lineItemQtyForDisplay(item);
    const price = parseFloat(String(item?.['Unit Price'] ?? item?.['Price'] ?? item?.['Amount'] ?? 0));
    subtotal += qty * price;
  }
  const dr = clampUnitRate(req.discount_rate !== undefined && req.discount_rate !== null ? Number(req.discount_rate) : 0);
  const tr = clampUnitRate(req.tax_rate !== undefined && req.tax_rate !== null ? Number(req.tax_rate) : 0.18);
  const subtotalRounded = roundMoney2(subtotal);
  const discountAmount = roundMoney2(subtotalRounded * dr);
  const taxableBase = roundMoney2(subtotalRounded - discountAmount);
  const taxAmount = roundMoney2(taxableBase * tr);
  const total = roundMoney2(taxableBase + taxAmount);
  return {
    subtotal: subtotalRounded,
    discountRate: dr,
    discountAmount,
    taxableBase,
    taxRate: tr,
    taxAmount,
    total,
  };
}

const PR_COLUMNS = [
  'Item',
  'Quantity',
  'Unit Price',
  'Cost Center',
  'Request to be delivered on',
  REMARKS_LINE_COL,
];

/** One supplier for the whole PR; falls back to legacy per-line "Suggested Supplier". */
function prSuggestedSupplierDisplay(request: Pick<WorkflowRequest, 'suggested_supplier' | 'line_items'>): string {
  const docLevel = (request.suggested_supplier ?? '').toString().trim();
  if (docLevel) return docLevel;
  for (const item of request.line_items || []) {
    const v = String(
      item?.['Suggested Supplier'] ?? item?.['suggested supplier'] ?? item?.Supplier ?? ''
    ).trim();
    if (v) return v;
  }
  return '';
}

const PO_COLUMNS = [
  'Item',
  'Final Supplier',
  'Quantity',
  'Unit Price',
  'Cost Center',
  'Request to be delivered on',
  REMARKS_LINE_COL,
];

/** Stock / spare requisition line grid (No is the row index in the UI table). */
const SR_COLUMNS = [
  'Item',
  'Unit Price',
  'Purchase From (Supplier)',
  'Min quantity',
  'Max quantity',
  'Spare for (Location)',
  'Reason',
];

const isProcurementNumericGridColumn = (col: string) => {
  const c = col.trim().toLowerCase();
  return (
    c === 'quantity' ||
    c === 'unit price' ||
    c === 'price' ||
    c === 'amount' ||
    c === 'min quantity' ||
    c === 'max quantity'
  );
};

/** Only these use min=0 on number inputs; unit price / amount may be negative (line discount). */
const isProcurementQuantityGridColumn = (col: string) => {
  const c = col.trim().toLowerCase();
  return c === 'quantity' || c === 'min quantity' || c === 'max quantity';
};

const isProcurementLineItemDateColumn = (col: string) => {
  const c = col.trim().toLowerCase();
  return c === 'delivery date' || c === 'request to be delivered on';
};

/** Normalize stored line-item dates to `YYYY-MM-DD` for `<input type="date">`. */
function htmlDateValueFromStored(raw: unknown): string {
  const s = String(raw ?? '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const head = s.slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(head) && (s.length === 10 || s[10] === 'T' || s[10] === ' ')) return head;
  const m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/.exec(s);
  if (m) {
    const p1 = parseInt(m[1], 10);
    const p2 = parseInt(m[2], 10);
    const y = parseInt(m[3], 10);
    let day: number;
    let month: number;
    if (p1 > 12) {
      day = p1;
      month = p2;
    } else if (p2 > 12) {
      month = p1;
      day = p2;
    } else {
      day = p1;
      month = p2;
    }
    if (y >= 1900 && y <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const d = new Date(y, month - 1, day);
      if (!Number.isNaN(d.getTime()) && d.getFullYear() === y && d.getMonth() === month - 1 && d.getDate() === day) {
        return `${y}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      }
    }
  }
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }
  return '';
}

function normalizeLineItemsForDateInputs(items: any[], columns: string[]): any[] {
  if (!Array.isArray(items) || items.length === 0) return items;
  const dateCols = columns.filter((col) => isProcurementLineItemDateColumn(col));
  if (dateCols.length === 0) return items;
  return items.map((it) => {
    if (!it || typeof it !== 'object') return it;
    const next = { ...it };
    for (const col of dateCols) {
      const cur = next[col];
      if (cur != null && String(cur).trim() !== '') {
        next[col] = htmlDateValueFromStored(cur);
      }
    }
    return next;
  });
}

const lineItemRemarksDisplay = (item: any) =>
  String(item?.[REMARKS_LINE_COL] ?? item?.['Remarks (Purpose)'] ?? '').trim();

const mergeLineItemRemarksWrite = (item: any, value: string) => {
  const next = { ...item, [REMARKS_LINE_COL]: value };
  delete (next as any)['Remarks (Purpose)'];
  return next;
};

/** Multi-line notes per line item (stored in JSON; not a template column). */
const LINE_ITEM_REMARKS_KEY = '_lineRemarks';

/** Signature images are stored as data URLs in NVARCHAR(MAX). Very short strings are usually legacy truncated rows. */
function isSignatureImageDataUrl(s: string | undefined | null): boolean {
  if (!s || typeof s !== 'string') return false;
  const t = s.trim();
  if (!/^data:image\/(png|jpeg|jpg);base64,/i.test(t)) return false;
  return t.length >= 80;
}

function formatSignatureProofTimestamp(iso: string | undefined | null): string {
  if (!iso) return '';
  try {
    return formatDateTimeMYT(iso);
  } catch {
    return '';
  }
}

function hasRequesterSignatureProof(req: WorkflowRequest): boolean {
  return !!(req.requester_signed_at || isSignatureImageDataUrl(req.requester_signature));
}

/** PR, PO, and procurement invoice flows require a drawn signature on each approval (stored for PDFs). */
function requiresProcurementApproverSignaturePad(req: WorkflowRequest): boolean {
  if (req.category !== 'procurement') return false;
  if (isPRRequest(req) || isPO_Only(req) || isSRRequest(req)) return true;
  return req.template_name.toLowerCase().includes('invoice');
}

const SignaturePad = ({
  onSave,
  onClear,
  value,
  savedSignature,
  onUseSaved,
  onSaveDefault,
  onClearSaved,
}: {
  onSave: (data: string) => void;
  onClear: () => void;
  /** Current signature image to display (data URL). */
  value?: string | null;
  savedSignature?: string | null;
  onUseSaved?: () => void;
  onSaveDefault?: () => void;
  onClearSaved?: () => void;
}) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const uploadInputRef = useRef<HTMLInputElement>(null);
  const isDrawingRef = useRef(false);

  const drawSignatureImageToCanvas = (img: HTMLImageElement, canvas: HTMLCanvasElement, ctx: CanvasRenderingContext2D) => {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    // Fit image into the signature box with safe padding while preserving aspect ratio.
    const pad = 12;
    const maxW = Math.max(1, canvas.width - pad * 2);
    const maxH = Math.max(1, canvas.height - pad * 2);
    const iw = img.naturalWidth || 1;
    const ih = img.naturalHeight || 1;
    let w = maxW;
    let h = (w * ih) / iw;
    if (h > maxH) {
      h = maxH;
      w = (h * iw) / ih;
    }
    const x = (canvas.width - w) / 2;
    const y = (canvas.height - h) / 2;
    ctx.drawImage(img, x, y, w, h);
  };

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 3;
    ctx.lineCap = 'round';
  }, []);

  // When caller sets a signature value (e.g. "Use saved"), render it into the canvas.
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const v = String(value ?? '').trim();
    if (!v) {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      return;
    }
    const img = new Image();
    img.onload = () => {
      drawSignatureImageToCanvas(img, canvas, ctx);
    };
    img.src = v;
  }, [value]);

  const getCanvasCoords = (e: React.MouseEvent | React.TouchEvent, canvas: HTMLCanvasElement) => {
    const rect = canvas.getBoundingClientRect();
    let clientX: number;
    let clientY: number;
    if ('touches' in e && e.touches.length > 0) {
      clientX = e.touches[0].clientX;
      clientY = e.touches[0].clientY;
    } else {
      const me = e as React.MouseEvent;
      clientX = me.clientX;
      clientY = me.clientY;
    }
    const rw = rect.width || 1;
    const rh = rect.height || 1;
    const x = ((clientX - rect.left) / rw) * canvas.width;
    const y = ((clientY - rect.top) / rh) * canvas.height;
    return { x, y };
  };

  const startDrawing = (e: React.MouseEvent | React.TouchEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    isDrawingRef.current = true;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const { x, y } = getCanvasCoords(e, canvas);
    ctx.beginPath();
    ctx.moveTo(x, y);
  };

  const stopDrawing = () => {
    isDrawingRef.current = false;
    const canvas = canvasRef.current;
    if (canvas) {
      onSave(canvas.toDataURL('image/png'));
    }
  };

  const draw = (e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawingRef.current) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const { x, y } = getCanvasCoords(e, canvas);
    ctx.lineTo(x, y);
    ctx.stroke();
  };

  const clear = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    onClear();
  };

  const openUploadPicker = () => {
    uploadInputRef.current?.click();
  };

  const handleUploadSignature = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    e.target.value = '';
    if (!f) return;
    if (!f.type.startsWith('image/')) {
      toast.error('Please upload an image file.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result || '');
      if (!dataUrl) return;
      const canvas = canvasRef.current;
      if (!canvas) return;
      const ctx = canvas.getContext('2d');
      if (!ctx) return;
      const img = new Image();
      img.onload = () => {
        drawSignatureImageToCanvas(img, canvas, ctx);
        onSave(canvas.toDataURL('image/png'));
        toast.success('Signature uploaded.');
      };
      img.onerror = () => toast.error('Failed to read uploaded signature image.');
      img.src = dataUrl;
    };
    reader.onerror = () => toast.error('Failed to read uploaded signature image.');
    reader.readAsDataURL(f);
  };

  return (
    <div className="space-y-2">
      <div className="border border-zinc-300 rounded-lg overflow-hidden bg-white">
        <canvas
          ref={canvasRef}
          width={1200}
          height={480}
          className="w-full h-100 cursor-crosshair touch-none"
          onMouseDown={startDrawing}
          onMouseMove={draw}
          onMouseUp={stopDrawing}
          onMouseOut={stopDrawing}
          onTouchStart={startDrawing}
          onTouchMove={draw}
          onTouchEnd={stopDrawing}
        />
      </div>
      <div className="flex flex-wrap items-center gap-2">
        <input
          ref={uploadInputRef}
          type="file"
          accept="image/png,image/jpeg,image/jpg,image/webp"
          className="hidden"
          onChange={handleUploadSignature}
        />
        <button
          type="button"
          onClick={clear}
          className="text-xs text-zinc-500 hover:text-zinc-700 font-medium flex items-center gap-1"
        >
          <RotateCcw className="w-3 h-3" />
          Cleared
        </button>
        <button
          type="button"
          onClick={openUploadPicker}
          className="text-xs font-medium flex items-center gap-1 px-2 py-1 rounded-md border border-zinc-200 text-zinc-600 hover:bg-zinc-50 hover:text-zinc-800"
          title="Upload a signature image"
        >
          <Upload className="w-3 h-3" />
          Upload
        </button>
        {onUseSaved ? (
          <button
            type="button"
            onClick={onUseSaved}
            disabled={!savedSignature}
            className={cn(
              "text-xs font-medium flex items-center gap-1 px-2 py-1 rounded-md border",
              savedSignature
                ? "border-zinc-200 text-zinc-600 hover:bg-zinc-50 hover:text-zinc-800"
                : "border-zinc-100 text-zinc-300 cursor-not-allowed"
            )}
            title={savedSignature ? "Use your saved signature" : "No saved signature yet"}
          >
            Use saved
          </button>
        ) : null}
        {onSaveDefault ? (
          <button
            type="button"
            onClick={onSaveDefault}
            className="text-xs font-medium flex items-center gap-1 px-2 py-1 rounded-md border border-indigo-200 text-indigo-700 hover:bg-indigo-50"
            title="Save this signature for future use"
          >
            Save as default
          </button>
        ) : null}
        {onClearSaved ? (
          <button
            type="button"
            onClick={onClearSaved}
            disabled={!savedSignature}
            className={cn(
              "text-xs font-medium flex items-center gap-1 px-2 py-1 rounded-md border",
              savedSignature
                ? "border-amber-200 text-amber-800 hover:bg-amber-50"
                : "border-zinc-100 text-zinc-300 cursor-not-allowed"
            )}
            title="Remove saved signature from server"
          >
            Remove saved
          </button>
        ) : null}
      </div>
    </div>
  );
};

const isPOTemplate = (template: { name: string; category: string }) => {
  const n = template.name.toLowerCase();
  return template.category === 'procurement' && (n.includes('purchase order') || n.includes('po'));
};

const isSRTemplate = (template: { name: string; category: string }) => {
  if (template.category !== 'procurement') return false;
  const n = template.name.toLowerCase();
  return n.includes('stock requisition') || (n.includes('stock') && n.includes('requisition'));
};

/** Purchase Request template (excludes PO and stock requisition templates). */
const isPRTemplate = (template: { name: string; category: string }) => {
  if (template.category !== 'procurement') return false;
  if (isPOTemplate(template)) return false;
  if (isSRTemplate(template)) return false;
  const n = template.name.toLowerCase();
  return n.includes('purchase request') || n.includes('pr');
};

/** Purchase Request instance (not PO). */
const isPRRequest = (request: WorkflowRequest) => {
  if (!request || request.category !== 'procurement') return false;
  if (isPO_Only(request)) return false;
  const n = request.template_name.toLowerCase();
  return n.includes('purchase request') || n.includes('pr');
};

const isPR_Only = (item: any) => {
  if (!item) return false;
  const name = (item.template_name || item.name || '').toLowerCase();
  return item.category === 'procurement' && (name.includes('purchase request') || name.includes('pr')) && !name.includes('order') && !name.includes('po');
};

const isPO_Only = (item: any) => {
  if (!item) return false;
  const name = (item.template_name || item.name || '').toLowerCase();
  return item.category === 'procurement' && (name.includes('purchase order') || name.includes('po'));
};

const isSR_Only = (item: any) => {
  if (!item) return false;
  if (item.category !== 'procurement' || isPO_Only(item)) return false;
  const name = (item.template_name || item.name || '').toLowerCase();
  return name.includes('stock requisition') || (name.includes('stock') && name.includes('requisition'));
};

const isSRRequest = (request: WorkflowRequest) => {
  if (!request || request.category !== 'procurement' || isPO_Only(request)) return false;
  const n = request.template_name.toLowerCase();
  return n.includes('stock requisition') || (n.includes('stock') && n.includes('requisition'));
};

/** Hide free-text request details (formerly justification) for procurement PR/PO/SR only. */
const isProcurementPRorPORequest = (req: WorkflowRequest | null | undefined) =>
  !!req && req.category === 'procurement' && (isPRRequest(req) || isPO_Only(req) || isSRRequest(req));

/** List/export "Supplier" column: PR/SR suggested supplier; PO finalized supplier per line. */
function procurementSupplierColumnDisplay(request: WorkflowRequest): string {
  if (isPO_Only(request)) {
    const suppliers = new Set<string>();
    for (const item of request.line_items || []) {
      const v = String((item as Record<string, unknown>)?.['Final Supplier'] ?? '').trim();
      if (v) suppliers.add(v);
    }
    if (suppliers.size === 0) return '-';
    return [...suppliers].join(', ');
  }
  const pr = prSuggestedSupplierDisplay(request);
  return pr ? pr : '-';
}

const procurementGridColumns = (req: WorkflowRequest) => {
  if (isPRRequest(req)) return PR_COLUMNS;
  if (isSRRequest(req)) return SR_COLUMNS;
  if (isPO_Only(req)) return PO_COLUMNS;
  return req.table_columns || [];
};

const procurementRowShowsLineTotal = (req: WorkflowRequest) =>
  isPRRequest(req) || isSRRequest(req) || isPO_Only(req);

const parseDepartmentCsv = (value: string | undefined) =>
  String(value || '')
    .split(',')
    .map((x) => x.trim().toLowerCase())
    .filter(Boolean);

const departmentsOverlap = (left: string | undefined, right: string | undefined) => {
  const lset = new Set(parseDepartmentCsv(left));
  const rightList = parseDepartmentCsv(right);
  if (lset.size === 0 || rightList.length === 0) return false;
  return rightList.some((d) => lset.has(d));
};

/** Must match server `PR_SIGN_ON_BEHALF_USERNAMES` (display names, comma-separated). */
const parsePrSignOnBehalfUsernames = () => {
  const raw = (import.meta as any).env?.VITE_PR_SIGN_ON_BEHALF_USERNAMES as string | undefined;
  const s = (raw && String(raw).trim()) || 'Gracelyn Tong';
  return s
    .split(',')
    .map((x) => x.trim().toLowerCase())
    .filter(Boolean);
};

const isPrSignOnBehalfUser = (user: User | null | undefined) => {
  const n = String(user?.username || '')
    .trim()
    .toLowerCase();
  if (!n) return false;
  return parsePrSignOnBehalfUsernames().includes(n);
};

const getColumns = (item: any) => {
  if (!item) return [];
  if (isPO_Only(item)) return PO_COLUMNS;
  if (isSR_Only(item)) return SR_COLUMNS;
  if (isPR_Only(item)) return PR_COLUMNS;
  return item.table_columns || [];
};

/** Matches server approval rules (SOM = cross-department for PO SOM step). */
const userCanApproveWorkflowStep = (user: User, request: WorkflowRequest, currentStep: WorkflowStep) => {
  const role = currentStep.approverRole.toLowerCase();
  const userHasRole = user.roles?.some((r) => r.toLowerCase() === role);
  const isAdmin = user.roles?.some((r) => r.toLowerCase() === 'admin');
  const isDirector = user.roles?.some((r) => r.toLowerCase() === 'director') && (user.department || '').toLowerCase() === 'management';
  const isSom = user.roles?.some((r) => r.toLowerCase() === 'som') && (user.department || '').toLowerCase() === 'management';
  if (isAdmin || isDirector) return true;
  if (role === 'som' && isSom) return true;
  const ent = (request.entity || '').trim().toUpperCase();
  if (
    isPrSignOnBehalfUser(user) &&
    isPRRequest(request) &&
    role === 'approver'
  )
    return true;
  if (
    ent === 'GCCM' &&
    role === 'approver' &&
    departmentsOverlap(user.department, request.department)
  )
    return true;
  if (userHasRole && departmentsOverlap(user.department, request.department)) return true;
  return false;
};

/** API stores workflow request status in lowercase (pending | approved | rejected). */
const normalizeWorkflowRequestStatus = (s: string | undefined) =>
  (s ?? '').toString().trim().toLowerCase();

/** All approval steps completed — server sets `status` to approved only after the final step. */
const isWorkflowRequestFullyApproved = (r: Pick<WorkflowRequest, 'status'>) =>
  normalizeWorkflowRequestStatus(r.status) === 'approved';

const isWorkflowRequestPending = (r: Pick<WorkflowRequest, 'status'>) =>
  normalizeWorkflowRequestStatus(r.status) === 'pending';

const isWorkflowRequestRejected = (r: Pick<WorkflowRequest, 'status'>) =>
  normalizeWorkflowRequestStatus(r.status) === 'rejected';

const isWorkflowRequestCancelled = (r: Pick<WorkflowRequest, 'status'>) =>
  normalizeWorkflowRequestStatus(r.status) === 'cancelled';

function purchasingCancelApprovalForRequest(req: WorkflowRequest | null | undefined): RequestApproval | undefined {
  if (!req) return undefined;
  const list = (req.approvals || []).filter((a) => normalizeWorkflowRequestStatus(a.status) === 'cancelled');
  if (list.length === 0) return undefined;
  // Prefer purchasing_final snapshot if present; else last cancelled approval.
  const byPurchasing =
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (list as any).findLast?.((a: RequestApproval) => String(a.approver_role_snapshot || '').toLowerCase() === 'purchasing_final') ??
    undefined;
  return byPurchasing || list[list.length - 1];
}

function requesterCancelApprovalForRequest(req: WorkflowRequest | null | undefined): RequestApproval | undefined {
  if (!req) return undefined;
  const list = (req.approvals || []).filter((a) => normalizeWorkflowRequestStatus(a.status) === 'cancelled');
  if (list.length === 0) return undefined;
  // Prefer requester cancellation record when present.
  const byRequester =
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (list as any).findLast?.((a: RequestApproval) => String(a.approver_role_snapshot || '').toLowerCase() === 'requester_cancel') ??
    undefined;
  return byRequester;
}

function purchasingRejectApprovalForRequest(req: WorkflowRequest | null | undefined): RequestApproval | undefined {
  if (!req) return undefined;
  const list = (req.approvals || []).filter((a) => normalizeWorkflowRequestStatus(a.status) === 'rejected');
  if (list.length === 0) return undefined;
  // Prefer purchasing_final snapshot if present; else last rejected approval.
  const byPurchasing =
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (list as any).findLast?.((a: RequestApproval) => String(a.approver_role_snapshot || '').toLowerCase() === 'purchasing_final') ??
    undefined;
  return byPurchasing || list[list.length - 1];
}

const workflowStatusNormalizedBadgeClass = (n: string) => {
  if (n === 'approved') return 'bg-emerald-100 text-emerald-700 ring-1 ring-inset ring-emerald-200/70';
  if (n === 'rejected') return 'bg-red-100 text-red-700 ring-1 ring-inset ring-red-200/70';
  if (n === 'cancelled') return 'bg-rose-100 text-rose-700 ring-1 ring-inset ring-rose-200/70';
  return 'bg-amber-100 text-amber-700 ring-1 ring-inset ring-amber-200/70';
};

const workflowRequestStatusBadgeClass = (r: WorkflowRequest) =>
  workflowStatusNormalizedBadgeClass(normalizeWorkflowRequestStatus(r.status));

const workflowRequestStatusBadgeClassFromRawStatus = (status: string | null | undefined) =>
  workflowStatusNormalizedBadgeClass(normalizeWorkflowRequestStatus(status));

/** Status text for a linked PO row (uses PO-approved wording when approved). */
const formatLinkedPoWorkflowStatusLabel = (linkedStatus: string | null | undefined) => {
  const n = normalizeWorkflowRequestStatus(linkedStatus);
  if (n === 'approved') return 'PO approved';
  if (n === 'rejected') return 'Rejected';
  if (n === 'cancelled') return 'Cancelled';
  if (n === 'pending') return 'Pending';
  return linkedStatus?.trim() || '—';
};

const formatWorkflowRequestStatusLabel = (r: WorkflowRequest) => {
  const n = normalizeWorkflowRequestStatus(r.status);
  if (n === 'approved') {
    if (isPO_Only(r)) return 'PO approved';
    if (isPR_Only(r)) return 'PR approved';
    return 'Approved';
  }
  if (n === 'rejected') return 'Rejected';
  if (n === 'cancelled') return 'Cancelled';
  if (n === 'pending') return 'Pending';
  return r.status ?? '—';
};

const requestHasChosenApprover = (r: WorkflowRequest): boolean => {
  const id = r.assigned_approver_id;
  return id != null && Number.isFinite(Number(id)) && Number(id) > 0;
};

const chosenApproverNameLabel = (r: WorkflowRequest): string => {
  const n = String(r.assigned_approver_name ?? '').trim();
  return n || '—';
};

/** PR already has a PO request created from it (one PR → one PO). */
const prHasLinkedPoRequest = (r: WorkflowRequest) =>
  r.converted_po_request_id != null && Number(r.converted_po_request_id) > 0;

/** PO # column: PO documents show their own formatted id; PRs show the linked PO id when converted. */
const displayPoNumberInRequestTable = (r: WorkflowRequest) => {
  if (isPO_Only(r) && String(r.formatted_id ?? '').trim()) return toUpperSerial(String(r.formatted_id).trim());
  const linked = String(r.linked_po_formatted_id ?? '').trim();
  if (linked) return toUpperSerial(linked);
  return '—';
};

/** PO workflow status: for PO rows, same as document status; for PRs with a linked PO, linked PO status. */
const displayPoStatusLabelInRequestTable = (r: WorkflowRequest): string => {
  if (isPO_Only(r)) return formatWorkflowRequestStatusLabel(r);
  if (isPR_Only(r) && prHasLinkedPoRequest(r)) return formatLinkedPoWorkflowStatusLabel(r.linked_po_status);
  return '—';
};

const showPoStatusBadgeInRequestTable = (r: WorkflowRequest) => {
  if (isPO_Only(r)) return true;
  if (isPR_Only(r) && prHasLinkedPoRequest(r)) return !!String(r.linked_po_status ?? '').trim();
  return false;
};

/** Purchasing users only; used for PR → PO actions. */
const isPurchasingRole = (user: User) =>
  !!user.roles?.some((role) => role.toLowerCase() === 'purchasing');

const isAdminPermission = (user: User) => !!user.permissions?.includes('admin');

/** Show Convert to PO only for fully approved procurement PRs on the purchasing side, and only once per PR. */
const canShowConvertPRToPO = (r: WorkflowRequest, user: User) =>
  isWorkflowRequestFullyApproved(r) &&
  isPR_Only(r) &&
  isPurchasingRole(user) &&
  !prHasLinkedPoRequest(r);

/** Purchasing: approved PR tools in header (draft PDF may still run after a PO is linked). */
const canShowPurchasingApprovedPRHeaderActions = (r: WorkflowRequest, user: User) =>
  isWorkflowRequestFullyApproved(r) && isPR_Only(r) && isPurchasingRole(user);

/** Purchasing can still override an already-approved PR with reject/cancel. */
const canShowPurchasingFinalDecision = (r: WorkflowRequest, user: User) =>
  isWorkflowRequestFullyApproved(r) && isPR_Only(r) && (isPurchasingRole(user) || isAdminPermission(user));

const canRequesterCancelPendingRequest = (
  r: WorkflowRequest,
  user: User,
  approvals: RequestApproval[] | undefined
) => {
  if (!isWorkflowRequestPending(r)) return false;
  if (user.id !== r.requester_id) return false;
  const hasApproved = (approvals || []).some((a) => normalizeWorkflowRequestStatus(a.status) === 'approved');
  return !hasApproved;
};

const stepApproverRoleLowerAt = (steps: WorkflowStep[] | undefined, stepIndex: number): string =>
  (steps?.[stepIndex]?.approverRole ?? '').toString().trim().toLowerCase();

/** Prefer `approver_role_snapshot` from DB at sign-off; fall back to template step role. */
const approvalStepRoleLower = (a: RequestApproval, steps: WorkflowStep[] | undefined): string => {
  const snap = (a.approver_role_snapshot ?? '').toString().trim().toLowerCase();
  if (snap) return snap;
  return stepApproverRoleLowerAt(steps, a.step_index);
};

const workflowApprovedApprovals = (request: WorkflowRequest): RequestApproval[] => {
  const list = (request.approvals || []).filter((a) => normalizeWorkflowRequestStatus(a.status) === 'approved');
  // Always prefer the most recent approval record per step_index.
  // (Reject/resubmit cycles can leave multiple rows for a step; UI/PDF should show latest signature.)
  const seen = new Set<number>();
  const out: RequestApproval[] = [];
  for (let i = list.length - 1; i >= 0; i--) {
    const a = list[i];
    const idx = Number(a.step_index);
    if (!Number.isFinite(idx)) continue;
    if (seen.has(idx)) continue;
    seen.add(idx);
    out.push(a);
  }
  return out.reverse();
};

/**
 * SR form (three columns): map signatures to existing boxes — checker → middle, else approver → middle; som → right, else approver when middle is checker.
 */
const srSignatureColumnsForPdf = (
  request: WorkflowRequest
): { mid: RequestApproval | undefined; senior: RequestApproval | undefined } => {
  const steps = request.template_steps || [];
  const list = workflowApprovedApprovals(request);
  const byRole = (role: string) => [...list].reverse().find((a) => approvalStepRoleLower(a, steps) === role);
  const templateHasChecker = (steps || []).some((s) => (s.approverRole || '').toLowerCase() === 'checker');
  if (templateHasChecker) {
    return { mid: byRole('checker'), senior: byRole('som') ?? byRole('approver') };
  }
  return { mid: byRole('approver'), senior: byRole('som') };
};

const pdfDesignationLines = (
  doc: jsPDF,
  prefix: string,
  designation: string | null | undefined,
  maxW: number,
  fontSize = 7
) => {
  const d = (designation ?? '').toString().trim();
  const s = d ? `${prefix}${d}` : `${prefix}—`;
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(fontSize);
  return doc.splitTextToSize(s, maxW);
};

/** F-PU-003 style form PDF � body in {@link buildProcurementPrFormPdfDoc}. */
const printProcurementPRFormPdf = (request: WorkflowRequest) => {
  const doc = buildProcurementPrFormPdfDoc(request);
  return buildPdfPreview(doc, `PR_${displayRequestSerial(request)}.pdf`);
};

/** Landscape stock item requisition form — matches company template (header, grid, 3 signatures, note). */
const printProcurementSRFormPdf = (request: WorkflowRequest) => {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageW = 297;
  const pageH = 210;
  const left = 10;
  const right = pageW - 10;
  const tableLeft = left;
  const tableRight = right;
  /** Min table rows (blank lines) like paper form. */
  const SR_MIN_DATA_ROWS = 1;
  const colWidths = [8, 38, 20, 44, 16, 16, 26, 32, 77];
  const colXs: number[] = [];
  {
    let x = tableLeft;
    colWidths.forEach((w) => {
      colXs.push(x);
      x += w;
    });
    colXs.push(tableRight);
  }

  const padX = 1.2;
  const padY = 1.8;
  const headerFont = 7;
  const bodyFont = 6;
  const lineStepHeader = 3.0;
  const lineStepBody = 2.75;
  const minRowH = 6.5;
  const pageBottomSafe = pageH - 12;
  /** Same table start on every page so form chrome + column headers line up on continuations. */
  const tableTopY = 40;

  const getItemValue = (item: any, keys: string[]) => {
    const foundKey = Object.keys(item || {}).find((k) => keys.includes(k.toLowerCase()));
    if (!foundKey) return '';
    return String(item[foundKey] ?? '');
  };

  const lineItems = request.line_items || [];

  const headerLabels = [
    'No',
    'Item',
    'Unit Price',
    'Purchase From (Supplier)',
    'Min Quantity',
    'Max Quantity',
    'Spare for (Location)',
    'Reason',
  ];

  const wrap = (text: string, colIndex: number, fontSize: number) => {
    doc.setFontSize(fontSize);
    const maxW = Math.max(4, colWidths[colIndex] - padX * 2);
    return doc.splitTextToSize(text || '-', maxW);
  };

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(headerFont);
  const headerLineArrays = headerLabels.map((label, i) => wrap(label, i, headerFont));
  const headerRowH =
    Math.max(1, ...headerLineArrays.map((lines) => lines.length)) * lineStepHeader + padY * 2;

  type RowBlock = { kind: 'data'; cells: string[][]; h: number };

  const dataBlocks: RowBlock[] = [];
  const approvedApprovals = (request.approvals || []).filter((a) => a.status.toLowerCase() === 'approved');

  const curPfx = procurementCurrencyPrefix(request.currency);

  lineItems.forEach((item, idx) => {
    doc.setFontSize(bodyFont);
    const itemName = getItemValue(item, ['item', 'description']);
    const priceRaw = getItemValue(item, ['unit price', 'price', 'amount']);
    const priceNum = Number(priceRaw);
    const priceDisp =
      priceRaw !== '' && !Number.isNaN(priceNum) ? `${curPfx}${formatProcurementMoney(priceNum)}`.trim() : '-';
    const supplier = getItemValue(item, [
      'purchase from (supplier)',
      'purchase from',
      'supplier',
      'suggested supplier',
      'vendor',
    ]);
    const minQ = getItemValue(item, ['min quantity']);
    const maxQ = getItemValue(item, ['max quantity']);
    const cc = getItemValue(item, ['cost center', 'cost center account no.', 'cost centre']);
    const spareLoc = getItemValue(item, ['spare for (location)', 'spare for', 'location']);
    const reason = getItemValue(item, ['reason']);
    const lineOnlyRemarks = item && item[LINE_ITEM_REMARKS_KEY] ? String(item[LINE_ITEM_REMARKS_KEY]).trim() : '';
    const itemCellText = lineOnlyRemarks ? `${itemName}\n\n${lineOnlyRemarks}` : itemName;

    const cells: string[][] = [
      wrap(String(idx + 1), 0, bodyFont),
      wrap(itemCellText, 1, bodyFont),
      wrap(priceDisp, 2, bodyFont),
      wrap(supplier, 3, bodyFont),
      wrap(minQ || '-', 4, bodyFont),
      wrap(maxQ || '-', 5, bodyFont),
      wrap(cc || '-', 6, bodyFont),
      wrap(spareLoc, 7, bodyFont),
      wrap(reason || '-', 8, bodyFont),
    ];
    const linesPerCol = cells.map((lines) => lines.length);
    const h = Math.max(minRowH, Math.max(...linesPerCol, 1) * lineStepBody + padY * 2);
    dataBlocks.push({ kind: 'data', cells, h });
  });

  const padRows = Math.max(0, SR_MIN_DATA_ROWS - lineItems.length);
  for (let p = 0; p < padRows; p++) {
    const cells: string[][] = [
      wrap('', 0, bodyFont),
      wrap('', 1, bodyFont),
      wrap('', 2, bodyFont),
      wrap('', 3, bodyFont),
      wrap('', 4, bodyFont),
      wrap('', 5, bodyFont),
      wrap('', 6, bodyFont),
      wrap('', 7, bodyFont),
      wrap('', 8, bodyFont),
    ];
    dataBlocks.push({ kind: 'data', cells, h: minRowH });
  }

  const subsidiaryLabel = entityLegalDisplayName(request.entity);
  const formRef = displayRequestSerial(request);

  const drawPageChrome = (tableTop: number) => {
    doc.setLineWidth(0.3);
    doc.rect(left, 8, right - left, pageH - 16);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.text('STOCK ITEM REQUISITION FORM', pageW / 2, 14, { align: 'center' });
    doc.setFontSize(8);
    doc.text('Department:', left + 2, 22);
    doc.setFont('helvetica', 'normal');
    doc.text(request.department || '-', left + 30, 22);
    doc.setFont('helvetica', 'bold');
    doc.text('Section:', left + 2, 26);
    doc.setFont('helvetica', 'normal');
    doc.text(request.section || '-', left + 30, 26);
    doc.setFont('helvetica', 'bold');
    doc.text('Date:', left + 2, 29);
    doc.setFont('helvetica', 'normal');
    doc.text(formatDateMYT(request.created_at), left + 20, 29);
    doc.setFont('helvetica', 'bold');
    doc.text('Subsidiary:', tableRight - 2, 22, { align: 'right' });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    const subMaxW = 115;
    const subLines = doc.splitTextToSize(subsidiaryLabel || '-', subMaxW);
    let subY = 25;
    subLines.forEach((ln) => {
      doc.text(ln, tableRight - 2, subY, { align: 'right' });
      subY += 3.6;
    });
    doc.setFontSize(7);
    doc.setTextColor(90, 90, 90);
    doc.text(`Ref: ${formRef}`, left + 2, Math.max(33, subY + 1));
    doc.setTextColor(0, 0, 0);
    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.line(tableLeft, tableTop, tableRight, tableTop);
  };

  const drawVerticalRules = (y0: number, y1: number) => {
    colXs.forEach((x) => doc.line(x, y0, x, y1));
  };

  const drawHeaderRow = (yTop: number) => {
    const yBot = yTop + headerRowH;
    doc.line(tableLeft, yBot, tableRight, yBot);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(headerFont);
    headerLineArrays.forEach((lines, i) => {
      let y = yTop + padY + lineStepHeader * 0.85;
      lines.forEach((line) => {
        doc.text(line, colXs[i] + padX, y);
        y += lineStepHeader;
      });
    });
    return yBot;
  };

  const drawDataRow = (yTop: number, rowH: number, cells: string[][]) => {
    const yBot = yTop + rowH;
    doc.line(tableLeft, yBot, tableRight, yBot);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(bodyFont);
    cells.forEach((lines, i) => {
      let y = yTop + padY + lineStepBody * 0.85;
      lines.forEach((line) => {
        doc.text(line, colXs[i] + padX, y);
        y += lineStepBody;
      });
    });
    return yBot;
  };

  let tableTop = tableTopY;
  let y = tableTop;
  drawPageChrome(tableTop);
  let headerBottom = drawHeaderRow(y);
  y = headerBottom;
  drawVerticalRules(tableTop, headerBottom);

  let blockIndex = 0;
  while (blockIndex < dataBlocks.length) {
    const block = dataBlocks[blockIndex];
    if (y + block.h > pageBottomSafe) {
      doc.addPage('a4', 'l');
      tableTop = tableTopY;
      y = tableTop;
      drawPageChrome(tableTop);
      headerBottom = drawHeaderRow(y);
      y = headerBottom;
      drawVerticalRules(tableTop, y);
    }
    y = drawDataRow(y, block.h, block.cells);
    drawVerticalRules(tableTop, y);
    blockIndex += 1;
  }

  const tableBottom = y;
  drawVerticalRules(tableTop, tableBottom);

  const srNoteText =
    "Note: Stock Item Requisition Form is just for those part that need to be reorder regularly once it's reach minimum quantity or those critical machine parts.";
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  const discMaxW = tableRight - tableLeft - 4;
  const discLines = doc.splitTextToSize(srNoteText, discMaxW);
  const discLineH = 3.35;
  const discBlockH = discLines.length * discLineH + 2;

  const signatureBlockH = 48;
  const signatureBoxInnerH = 44;
  const footerGap = 3;
  const footerTotalH = signatureBlockH + footerGap + discBlockH;
  let signTop = tableBottom + 4;
  if (signTop + footerTotalH > pageBottomSafe) {
    doc.addPage('a4', 'l');
    drawPageChrome(tableTopY);
    signTop = tableTopY + 4;
  }
  const signBottom = signTop + signatureBoxInnerH;
  const { mid: opManagerApproval, senior: seniorApproval } = srSignatureColumnsForPdf(request);

  const addSignatureImageFit = (dataUrl: string, xMm: number, yTopMm: number, maxW: number, maxH: number) => {
    if (!dataUrl) return;
    const lower = dataUrl.toLowerCase();
    const fmtGuess: 'PNG' | 'JPEG' =
      lower.includes('image/jpeg') || lower.includes('image/jpg') ? 'JPEG' : 'PNG';
    const fitToBox = (iw: number, ih: number) => {
      let w = maxW;
      let h = (w * ih) / iw;
      if (h > maxH) {
        h = maxH;
        w = (h * iw) / ih;
      }
      return { w, h };
    };
    try {
      const { width: iw, height: ih } = doc.getImageProperties(dataUrl);
      if (iw > 0 && ih > 0) {
        const { w, h } = fitToBox(iw, ih);
        try {
          doc.addImage(dataUrl, fmtGuess, xMm, yTopMm, w, h);
        } catch {
          doc.addImage(dataUrl, fmtGuess === 'PNG' ? 'JPEG' : 'PNG', xMm, yTopMm, w, h);
        }
        return;
      }
    } catch {
      /* fall through */
    }
    const { w, h } = fitToBox(400, 150);
    for (const fmt of ['PNG', 'JPEG'] as const) {
      try {
        doc.addImage(dataUrl, fmt, xMm, yTopMm, w, h);
        return;
      } catch {
        /* try next format */
      }
    }
  };

  const colW = (right - left) / 3;
  const xDiv1 = left + colW;
  const xDiv2 = left + 2 * colW;
  doc.rect(left, signTop, right - left, signBottom - signTop);
  doc.line(xDiv1, signTop, xDiv1, signBottom);
  doc.line(xDiv2, signTop, xDiv2, signBottom);

  const sigPad = 2;
  const sigMaxW = colW - sigPad * 2;
  const sigMaxH = 16;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  let ly = signTop + 4;
  doc.text('Requested By:', left + sigPad, ly);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(6.5);
  const lblMid = doc.splitTextToSize('Approved By Operation Manager:', sigMaxW);
  let my = signTop + 4;
  lblMid.forEach((ln) => {
    doc.text(ln, xDiv1 + sigPad, my);
    my += 3.2;
  });
  const lblSen = doc.splitTextToSize('Approved By Senior Operation Manager', sigMaxW);
  let sy = signTop + 4;
  lblSen.forEach((ln) => {
    doc.text(ln, xDiv2 + sigPad, sy);
    sy += 3.2;
  });
  const sigImgY = Math.max(signTop + 14, my, sy) + 1;

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  const reqSigUrl = isSignatureImageDataUrl(request.requester_signature) ? request.requester_signature! : '';
  if (reqSigUrl) {
    addSignatureImageFit(reqSigUrl, left + sigPad, sigImgY, sigMaxW, sigMaxH);
  } else if (request.requester_signed_at) {
    doc.text('Electronic signature', left + sigPad, sigImgY + 4);
    doc.text(formatSignatureProofTimestamp(request.requester_signed_at), left + sigPad, sigImgY + 9);
  }

  const opSigUrl =
    opManagerApproval?.approver_signature && isSignatureImageDataUrl(opManagerApproval.approver_signature)
      ? opManagerApproval.approver_signature!
      : '';
  if (opSigUrl) {
    addSignatureImageFit(opSigUrl, xDiv1 + sigPad, sigImgY, sigMaxW, sigMaxH);
  } else if (opManagerApproval && opManagerApproval.status.toLowerCase() === 'approved') {
    doc.text('Electronic signature', xDiv1 + sigPad, sigImgY + 4);
    doc.text(formatSignatureProofTimestamp(opManagerApproval.created_at), xDiv1 + sigPad, sigImgY + 9);
  }

  const senSigUrl =
    seniorApproval?.approver_signature && isSignatureImageDataUrl(seniorApproval.approver_signature)
      ? seniorApproval.approver_signature!
      : '';
  if (senSigUrl) {
    addSignatureImageFit(senSigUrl, xDiv2 + sigPad, sigImgY, sigMaxW, sigMaxH);
  } else if (seniorApproval && seniorApproval.status.toLowerCase() === 'approved') {
    doc.text('Electronic signature', xDiv2 + sigPad, sigImgY + 4);
    doc.text(formatSignatureProofTimestamp(seniorApproval.created_at), xDiv2 + sigPad, sigImgY + 9);
  }

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.text(`Name: ${request.requester_name || ''}`, left + sigPad, signBottom - 13);
  let ySrDes = signBottom - 8;
  pdfDesignationLines(doc, 'Designation: ', request.requester_designation, sigMaxW, 6.2).forEach((ln) => {
    doc.text(ln, left + sigPad, ySrDes);
    ySrDes += 3.1;
  });
  doc.text(`Name: ${opManagerApproval?.approver_name || ''}`, xDiv1 + sigPad, signBottom - 13);
  ySrDes = signBottom - 8;
  pdfDesignationLines(doc, 'Designation: ', opManagerApproval?.approver_designation ?? null, sigMaxW, 6.2).forEach((ln) => {
    doc.text(ln, xDiv1 + sigPad, ySrDes);
    ySrDes += 3.1;
  });
  doc.text(`Name: ${seniorApproval?.approver_name || ''}`, xDiv2 + sigPad, signBottom - 13);
  ySrDes = signBottom - 8;
  pdfDesignationLines(doc, 'Designation: ', seniorApproval?.approver_designation ?? null, sigMaxW, 6.2).forEach((ln) => {
    doc.text(ln, xDiv2 + sigPad, ySrDes);
    ySrDes += 3.1;
  });

  let noteY = signBottom + footerGap;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(7);
  discLines.forEach((ln) => {
    doc.text(ln, left + 2, noteY);
    noteY += discLineH;
  });
  doc.setFont('helvetica', 'normal');

  return buildPdfPreview(doc, `SR_${displayRequestSerial(request)}.pdf`);
};

/** Standard report-style PDF (line items + approvals) — not the F-PU-003 form. */
const printWorkflowRequestReportPdf = (request: WorkflowRequest) => {
  const doc = new jsPDF();
  const margin = 20;
  let y = 20;

  doc.setFontSize(20);
  doc.setFont('helvetica', 'bold');
  doc.text('PURCHASE REQUISITION', margin, y);
  y += 10;

  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  const entCode = request.entity?.trim() || '-';
  doc.text(`Entity: ${entityLegalDisplayName(request.entity)} (${entCode})`, margin, y);
  doc.text(`Date: ${formatDateMYT(request.created_at)}`, 150, y);
  y += 10;

  doc.setDrawColor(200);
  doc.line(margin, y, 190, y);
  y += 10;

  doc.setFont('helvetica', 'bold');
  doc.text('Request Information', margin, y);
  y += 7;
  doc.setFont('helvetica', 'normal');
  doc.text(`PR ID: ${displayRequestSerial(request)}`, margin, y);
  y += 5;
  doc.text(`Requester: ${request.requester_name}`, margin, y);
  y += 5;
  doc.text(`Department: ${request.department}`, margin, y);
  y += 5;
  doc.text(`Cost Center: ${request.cost_center || '-'}`, margin, y);
  y += 5;
  doc.text(`Title: ${request.title}`, margin, y);
  y += 10;
  if (isPRRequest(request)) {
    doc.text(`Suggested supplier: ${prSuggestedSupplierDisplay(request) || '-'}`, margin, y);
    y += 10;
  }

  if (!isPRRequest(request) && !isPO_Only(request) && !isSRRequest(request)) {
    doc.setFont('helvetica', 'bold');
    doc.text('Details:', margin, y);
    y += 7;
    doc.setFont('helvetica', 'normal');
    const detailsLines = doc.splitTextToSize(request.details || '', 170);
    doc.text(detailsLines, margin, y);
    y += (detailsLines.length * 5) + 10;
  }

  doc.setFont('helvetica', 'bold');
  doc.text('Line Items', margin, y);
  y += 7;

  const columns = getColumns(request);
  const headers = ['No', ...columns];
  if (procurementRowShowsLineTotal(request)) headers.push('Total');

  doc.setFillColor(240, 240, 240);
  doc.rect(margin, y - 5, 170, 7, 'F');
  doc.setFontSize(8);
  let x = margin;

  const colWidth = 170 / headers.length;
  headers.forEach((h) => {
    doc.text(h, x + 2, y);
    x += colWidth;
  });
  y += 7;

  doc.setFont('helvetica', 'normal');
  let subtotal = 0;
  request.line_items?.forEach((item, idx) => {
    if (y > 270) {
      doc.addPage();
      y = 20;
    }
    x = margin;
    doc.text(String(idx + 1), x + 2, y);
    x += colWidth;

    columns.forEach((col) => {
      let val = String(item[col] || '-');
      if (col === 'Unit Price' || col === 'Price' || col === 'Amount') {
        val = `${procurementCurrencyPrefix(request.currency)}${formatProcurementMoney(Number(item[col] || 0))}`.trim();
      }
      doc.text(val, x + 2, y, { maxWidth: colWidth - 4 });
      x += colWidth;
    });

    if (procurementRowShowsLineTotal(request)) {
      const qty = lineItemQtyForDisplay(item);
      const price = parseFloat(item['Unit Price'] || '0');
      const total = qty * price;
      subtotal += total;
      doc.text(`${procurementCurrencyPrefix(request.currency)}${formatProcurementMoney(total)}`.trim(), x + 2, y);
    }
    y += 7;
  });

  if (isPRRequest(request) || isPO_Only(request) || isSRRequest(request)) {
    y += 5;
    const m = procurementMoneyTotals(request);
    const curp = procurementCurrencyPrefix(request.currency);
    doc.setFont('helvetica', 'bold');
    doc.text(`Subtotal: ${curp}${formatProcurementMoney(m.subtotal)}`.trim(), 110, y);
    y += 5;
    doc.text(`Discount (${(m.discountRate * 100).toFixed(0)}%): ${curp}${formatProcurementMoney(m.discountAmount)}`.trim(), 110, y);
    y += 5;
    const tlab = procurementTaxLabelForEntity(request.entity);
    doc.text(`${tlab} (${(m.taxRate * 100).toFixed(0)}%): ${curp}${formatProcurementMoney(m.taxAmount)}`.trim(), 110, y);
    y += 7;
    doc.setFontSize(12);
    doc.text(`TOTAL: ${curp}${formatProcurementMoney(m.total)}`.trim(), 110, y);
    y += 15;
  }

  if (y > 240) {
    doc.addPage();
    y = 20;
  }
  doc.setFontSize(10);
  doc.setFont('helvetica', 'bold');
  doc.text('Approval History', margin, y);
  y += 10;

  request.template_steps.forEach((step, i) => {
    const approval = request.approvals?.find(a => a.step_index === i);
    doc.setFont('helvetica', 'bold');
    doc.text(`${step.label}:`, margin, y);
    doc.setFont('helvetica', 'normal');
    if (approval) {
      doc.text(`${approval.status.toUpperCase()} by ${approval.approver_name} on ${formatDateMYT(approval.created_at)}`, margin + 50, y);
      let stepH = 12;
      const proxyNm = String(approval.signed_by_name ?? '').trim();
      if (proxyNm) {
        doc.setFontSize(8);
        doc.setTextColor(67, 56, 202);
        doc.text(`Signed by proxy: ${proxyNm}`, margin + 50, y + 4);
        doc.setTextColor(0, 0, 0);
        doc.setFontSize(10);
        stepH = 16;
      }
      const asig = approval.approver_signature;
      if (isSignatureImageDataUrl(asig)) {
        try {
          doc.addImage(asig!, 'PNG', margin + 140, y - 8, 30, 10);
        } catch {
          /* ignore */
        }
      } else if (approval.status.toLowerCase() === 'approved') {
        doc.setFontSize(7);
        doc.text('E-signed', margin + 128, y - 2);
        doc.setFontSize(10);
      }
      y += stepH;
    } else {
      doc.text('Pending', margin + 50, y);
      y += 12;
    }
  });

  return buildPdfPreview(doc, `PR_${displayRequestSerial(request)}.pdf`);
};

/** Prefer `/approvals` payload (designation + signatures) when viewing details for this request. */
function mergeRequestApprovalsForPdf(
  request: WorkflowRequest,
  selectedId: number | undefined,
  detailApprovals: RequestApproval[]
): WorkflowRequest {
  if (detailApprovals.length > 0 && selectedId != null && request.id === selectedId) {
    return { ...request, approvals: detailApprovals };
  }
  return request;
}

function createWorkflowRequestPdfPreview(request: WorkflowRequest): { url: string; fileName: string; pdfDataUrl: string } {
  if (isPR_Only(request)) return printProcurementPRFormPdf(request);
  if (isSR_Only(request)) return printProcurementSRFormPdf(request);
  return printWorkflowRequestReportPdf(request);
}

/**
 * When the last workflow step is approved, persist the F-PU-003 PR / SR form PDF to the request’s
 * attachment folder (same as “View PR/SR form”). No-op for non–PR/SR or mid-chain approvals.
 */
async function persistPrSrFormPdfAfterWorkflowCompleted(
  requestId: number,
  requestBeforeTransition: WorkflowRequest,
  opts?: { skipFinalStepCheck?: boolean }
): Promise<void> {
  if (!isPR_Only(requestBeforeTransition) && !isSR_Only(requestBeforeTransition)) return;
  if (!opts?.skipFinalStepCheck) {
    const stepsLen = (requestBeforeTransition.template_steps || []).length;
    if (stepsLen === 0) return;
    if (requestBeforeTransition.current_step_index + 1 < stepsLen) return;
  }
  const ent = String(requestBeforeTransition.entity || '').trim();
  if (ent) api.setActiveEntity(ent);
  try {
    const appData = await api.request(`/api/workflow-requests/${requestId}/approvals`);
    const merged = mergeRequestApprovalsForPdf(requestBeforeTransition, requestId, appData);
    const preview = createWorkflowRequestPdfPreview(merged);
    await persistGeneratedProcurementFormPdf(requestId, preview.pdfDataUrl);
  } catch (e) {
    console.error('Failed to auto-archive procurement form PDF:', e);
  }
}

const FIXED_PR_STEPS = [{ id: 'step-1', label: 'Approver', approverRole: 'approver' }];

/** PO template definition (director step is omitted per request when total ≤ RM30k equivalent). */
const FIXED_PO_STEPS_FULL = [
  { id: 'po-step-1', label: 'Checker Verification', approverRole: 'checker' },
  { id: 'po-step-2', label: 'Final Approval', approverRole: 'som' },
  { id: 'po-step-3', label: 'Director Authorization (> RM30,000)', approverRole: 'director' },
];

/** SR template definition: fixed two-step approval chain. */
const FIXED_SR_STEPS = [
  { id: 'sr-step-1', label: 'HOD Approval', approverRole: 'approver' },
  { id: 'sr-step-2', label: 'Final Approval', approverRole: 'som' },
];

const CURRENCIES = ['MYR', 'SGD', 'EURO', 'GBP', 'USD', 'FCFA'];

/** UI: show stored currency only — no default code. */
function procurementCurrencyLabel(code: string | undefined | null): string {
  const s = String(code ?? '').trim();
  return s || '—';
}

/** Prefix for amounts in PDFs/lists when currency exists. */
function procurementCurrencyPrefix(code: string | undefined | null): string {
  const s = String(code ?? '').trim();
  return s ? `${s} ` : '';
}

/** One row from GET `/api/cost-centers` → pick list (stored value is always `code` for PR/SR grids). */
type CostCenterPickOption = {
  value: string;
  code: string;
  name: string;
  glAccount: string;
  searchLower: string;
};

function costCenterPickOptionsFromRows(rows: unknown): CostCenterPickOption[] {
  if (!Array.isArray(rows)) return [];
  const seen = new Set<string>();
  const out: CostCenterPickOption[] = [];
  for (const r of rows as { code?: unknown; name?: unknown; gl_account?: unknown; glAccount?: unknown }[]) {
    const code = String(r?.code ?? '').trim();
    if (!code || seen.has(code)) continue;
    seen.add(code);
    const name = String(r?.name ?? '').trim();
    const glAccount = String((r as any)?.gl_account ?? (r as any)?.glAccount ?? '').trim();
    const searchLower = `${code}\n${name}\n${glAccount}`.toLowerCase();
    out.push({ value: code, code, name, glAccount, searchLower });
  }
  out.sort((a, b) => a.code.localeCompare(b.code, undefined, { numeric: true }));
  return out;
}

/** Match a stored line cell to a cost_centers row (code, name, or "code — name" from older UIs). */
function catalogPickMatch(raw: string | undefined | null, options: CostCenterPickOption[]): CostCenterPickOption | undefined {
  const t = String(raw ?? '').trim();
  if (!t) return undefined;
  const lower = t.toLowerCase();
  const byCode = options.find((o) => o.code === t);
  if (byCode) return byCode;
  const byName = options.find((o) => o.name && o.name.toLowerCase() === lower);
  if (byName) return byName;
  for (const o of options) {
    const emDash = `${o.code} — `;
    const hyphen = `${o.code} - `;
    if (t.startsWith(emDash) || lower.startsWith(emDash.toLowerCase())) return o;
    if (t.startsWith(hyphen)) return o;
  }
  return undefined;
}

function formatCostCenterCatalogCell(raw: string | undefined | null, options: CostCenterPickOption[]): string {
  const t = String(raw ?? '').trim();
  if (!t) return '-';
  const m = catalogPickMatch(t, options);
  if (!m) return t;
  return m.name ? `${m.code} — ${m.name}` : m.code;
}

/** Persist only catalog `code` for Cost Center / Spare Location cells. */
function normalizeCatalogCodeValue(raw: unknown, options: CostCenterPickOption[]): string {
  const t = String(raw ?? '').trim();
  if (!t) return '';
  const m = catalogPickMatch(t, options);
  return m?.code || t;
}

function normalizeProcurementCatalogLineItems(
  items: any[],
  options: CostCenterPickOption[],
  forSR: boolean
): any[] {
  if (!Array.isArray(items) || items.length === 0) return [];
  return items.map((item) => {
    if (!item || typeof item !== 'object') return item;
    const next = { ...item };
    for (const col of Object.keys(next)) {
      if (isCostCenterGridColumn(col) || (forSR && isSpareLocationColumn(col))) {
        next[col] = normalizeCatalogCodeValue(next[col], options);
      }
    }
    return next;
  });
}

const normalizeGridColumnName = (col: string): string =>
  String(col || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');

const COST_CENTER_GRID_COLUMN_KEYS = new Set(['cost center', 'cost center account no.', 'cost centre']);
const SPARE_LOCATION_GRID_COLUMN_KEYS = new Set(['spare for (location)', 'spare for']);

const isCostCenterGridColumn = (col: string) => COST_CENTER_GRID_COLUMN_KEYS.has(normalizeGridColumnName(col));

/** Active rows from `dbo.cost_centers` via GET `/api/cost-centers`. */
function useCostCenterOptions(): {
  options: CostCenterPickOption[];
  loading: boolean;
  failed: boolean;
  fetchError: string | null;
} {
  const [activeEntity, setActiveEntity] = useState<string>(() => String(localStorage.getItem(FLOWMASTER_ENTITY_KEY) || '').trim().toUpperCase());
  const [opts, setOpts] = useState<CostCenterPickOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [failed, setFailed] = useState(false);
  const [fetchError, setFetchError] = useState<string | null>(null);

  useEffect(() => {
    const syncEntity = () => setActiveEntity(String(localStorage.getItem(FLOWMASTER_ENTITY_KEY) || '').trim().toUpperCase());
    window.addEventListener('storage', syncEntity);
    window.addEventListener(FLOWMASTER_ENTITY_CHANGED_EVENT, syncEntity as EventListener);
    return () => {
      window.removeEventListener('storage', syncEntity);
      window.removeEventListener(FLOWMASTER_ENTITY_CHANGED_EVENT, syncEntity as EventListener);
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        if (!activeEntity) {
          if (!cancelled) {
            setOpts([]);
            setFailed(false);
            setFetchError(null);
            setLoading(false);
          }
          return;
        }
        setFailed(false);
        setFetchError(null);
        setLoading(true);
        const rows = await api.request('/api/cost-centers');
        if (cancelled) return;
        setOpts(costCenterPickOptionsFromRows(rows));
      } catch (e: unknown) {
        if (!cancelled) {
          setOpts([]);
          setFailed(true);
          const msg = e instanceof Error ? e.message : String(e);
          setFetchError(msg);
          console.error("GET /api/cost-centers failed:", e);
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [activeEntity]);
  return { options: opts, loading, failed, fetchError };
}

const isSpareLocationColumn = (col: string) => SPARE_LOCATION_GRID_COLUMN_KEYS.has(normalizeGridColumnName(col));

const SearchableSelect = ({
  options,
  value,
  onChange,
  placeholder,
  loading,
  loadFailed,
  loadErrorHint,
}: {
  options: CostCenterPickOption[];
  value: string;
  onChange: (val: string) => void;
  placeholder: string;
  loading?: boolean;
  loadFailed?: boolean;
  /** Full error text from the API (shown when the catalog fails to load). */
  loadErrorHint?: string | null;
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const containerRef = useRef<HTMLDivElement>(null);

  const selected = useMemo(() => {
    const exact = options.find((o) => o.value === value);
    if (exact) return exact;
    return catalogPickMatch(value, options);
  }, [options, value]);
  const summaryText = useMemo(() => {
    if (loading) return 'Loading list…';
    if (loadFailed && options.length === 0) return 'Could not load list';
    if (selected) {
      const base = selected.name ? `${selected.code} — ${selected.name}` : selected.code;
      return selected.glAccount ? `${base} (G/L ${selected.glAccount})` : base;
    }
    if (value.trim()) return value;
    return '';
  }, [loading, loadFailed, options.length, selected, value]);

  const filteredOptions = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return options;
    return options.filter((opt) => opt.searchLower.includes(q));
  }, [options, search]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  useEffect(() => {
    if (!isOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setIsOpen(false);
        setSearch('');
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [isOpen]);

  const showPlaceholder = !summaryText;
  const canOpen = !loading;

  return (
    <div className="relative min-w-[12rem]" ref={containerRef}>
      <div
        onClick={() => canOpen && setIsOpen(!isOpen)}
        className={cn(
          'w-full min-h-[2.5rem] px-3 py-2 rounded-lg border bg-white flex items-center gap-2 transition-colors',
          canOpen ? 'border-zinc-300 cursor-pointer hover:border-zinc-400' : 'border-zinc-200 cursor-wait opacity-80'
        )}
      >
        <span
          className={cn(
            'flex-1 text-left text-sm leading-snug line-clamp-2',
            showPlaceholder ? 'text-zinc-400' : 'text-zinc-900'
          )}
        >
          {showPlaceholder ? placeholder : summaryText}
        </span>
        {value && !loading ? (
          <button
            type="button"
            className="shrink-0 p-0.5 rounded-md text-zinc-400 hover:text-zinc-700 hover:bg-zinc-100"
            aria-label="Clear"
            onClick={(e) => {
              e.stopPropagation();
              onChange('');
              setSearch('');
            }}
          >
            <X className="w-3.5 h-3.5" />
          </button>
        ) : null}
        <ChevronDown className={cn('w-4 h-4 shrink-0 text-zinc-400 transition-transform', isOpen && canOpen && 'rotate-180')} />
      </div>

      <AnimatePresence>
        {isOpen && canOpen && (
          <motion.div
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="absolute z-[100] left-0 right-0 mt-1 bg-white border border-zinc-200 rounded-xl shadow-xl max-h-80 overflow-hidden flex flex-col"
          >
            <div className="p-2 border-b border-zinc-100 flex items-center gap-2">
              <Search className="w-4 h-4 text-zinc-400 shrink-0 ml-1" />
              <input
                type="text"
                autoFocus
                className="flex-1 min-w-0 px-2 py-1.5 text-sm border border-zinc-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                placeholder="Search code or name…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                onClick={(e) => e.stopPropagation()}
              />
            </div>
            <div className="overflow-y-auto overscroll-contain py-1">
              {loadFailed && options.length === 0 ? (
                <div className="px-4 py-3 text-sm text-amber-800 bg-amber-50/90 space-y-2">
                  <p>The list could not be loaded (cost centers catalog). Refresh the page or ask an admin to check the server.</p>
                  {loadErrorHint ? (
                    <p className="text-xs text-amber-950/80 font-mono break-words whitespace-pre-wrap">{loadErrorHint}</p>
                  ) : null}
                </div>
              ) : filteredOptions.length > 0 ? (
                filteredOptions.map((opt) => (
                  <button
                    type="button"
                    key={opt.value}
                    onClick={() => {
                      onChange(opt.value);
                      setIsOpen(false);
                      setSearch('');
                    }}
                    className={cn(
                      'w-full text-left px-3 py-2.5 transition-colors border-b border-zinc-50 last:border-b-0',
                      value === opt.value ? 'bg-indigo-50' : 'hover:bg-indigo-50/60'
                    )}
                  >
                    <div className="text-sm font-semibold text-zinc-900 leading-snug">{opt.name || opt.code}</div>
                    <div className="text-[11px] text-zinc-500 font-mono mt-0.5 tracking-tight">
                      {opt.code}
                      {opt.glAccount ? ` · G/L ${opt.glAccount}` : ''}
                    </div>
                  </button>
                ))
              ) : (
                <div className="px-4 py-3 text-sm text-zinc-400">
                  {search.trim() ? 'No matching items' : 'No items in catalog'}
                </div>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
};

const WorkflowRequestCreator = ({ template, entity, onSuccess }: { template: Workflow, entity?: string, onSuccess: () => void }) => {
  const [title, setTitle] = useState('');
  const [details, setDetails] = useState('');
  const [currency, setCurrency] = useState('');
  const [section, setSection] = useState("");
  const [suggestedSupplier, setSuggestedSupplier] = useState('');
  const [lineItems, setLineItems] = useState<any[]>([]);
  const [attachments, setAttachments] = useState<{ name: string; type: string; data: string }[]>([]);
  const [signature, setSignature] = useState<string | null>(null);
  const [savedSignature, setSavedSignature] = useState<string | null>(null);
  const [savedSignatureLoading, setSavedSignatureLoading] = useState(false);
  const [loading, setLoading] = useState(false);
  const [taxRateInput, setTaxRateInput] = useState('');
  const [discountRateInput, setDiscountRateInput] = useState('');
  const [eligibleApprovers, setEligibleApprovers] = useState<{ id: number; username: string; department: string }[]>([]);
  const [assignedApproverId, setAssignedApproverId] = useState<number | ''>('');

  const isPR = isPRTemplate(template);
  const isPO = isPOTemplate(template);
  const isSR = isSRTemplate(template);
  const {
    options: costCenterOptions,
    loading: costCentersLoading,
    failed: costCentersLoadFailed,
    fetchError: costCentersFetchError,
  } = useCostCenterOptions();
  const columns = getColumns(template);
  const needsApproverPicker =
    (template.steps || []).some((s) => (s.approverRole || '').toLowerCase() === 'approver');

  useEffect(() => {
    if (!needsApproverPicker) {
      setEligibleApprovers([]);
      setAssignedApproverId('');
      return;
    }
    let cancelled = false;
    (async () => {
      try {
        const list = await api.request('/api/users/eligible-approvers');
        if (!cancelled && Array.isArray(list)) setEligibleApprovers(list);
      } catch {
        if (!cancelled) setEligibleApprovers([]);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [needsApproverPicker, entity]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setSavedSignatureLoading(true);
        const r = await api.request("/api/me/signature", { skipEntity: true });
        if (cancelled) return;
        setSavedSignature(r?.exists && typeof r?.dataUrl === "string" ? r.dataUrl : null);
      } catch {
        if (!cancelled) setSavedSignature(null);
      } finally {
        if (!cancelled) setSavedSignatureLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const useSavedSignature = async () => {
    try {
      setSavedSignatureLoading(true);
      const r = await api.request("/api/me/signature", { skipEntity: true });
      const next = r?.exists && typeof r?.dataUrl === "string" ? r.dataUrl : null;
      setSavedSignature(next);
      if (!next) {
        toast.error("No saved signature found.");
        return;
      }
      setSignature(next);
      toast.success("Using saved signature.");
    } catch (e: any) {
      toast.error(e?.message || "Failed to load saved signature");
    } finally {
      setSavedSignatureLoading(false);
    }
  };

  const saveSignatureAsDefault = async () => {
    if (!signature || !isSignatureImageDataUrl(signature)) {
      toast.error("Please draw or upload a signature first.");
      return;
    }
    try {
      setSavedSignatureLoading(true);
      await api.request("/api/me/signature", {
        method: "PUT",
        skipEntity: true,
        body: JSON.stringify({ dataUrl: signature }),
      });
      setSavedSignature(signature);
      toast.success("Saved signature.");
    } catch (e: any) {
      toast.error(e?.message || "Failed to save signature");
    } finally {
      setSavedSignatureLoading(false);
    }
  };

  const removeSavedSignature = async () => {
    try {
      setSavedSignatureLoading(true);
      await api.request("/api/me/signature", { method: "DELETE", skipEntity: true });
      setSavedSignature(null);
      toast.success("Removed saved signature.");
    } catch (e: any) {
      toast.error(e?.message || "Failed to remove saved signature");
    } finally {
      setSavedSignatureLoading(false);
    }
  };

  const addLineItem = () => {
    const newItem: any = { id: Math.random().toString(36).substr(2, 9), [LINE_ITEM_REMARKS_KEY]: '' };
    columns.forEach(col => newItem[col] = '');
    setLineItems([...lineItems, newItem]);
  };

  const removeLineItem = (index: number) => {
    setLineItems(lineItems.filter((_, i) => i !== index));
  };

  const updateLineItem = (index: number, col: string, value: string) => {
    const newItems = [...lineItems];
    if (col === REMARKS_LINE_COL) {
      newItems[index] = mergeLineItemRemarksWrite(newItems[index], value);
    } else {
      newItems[index][col] = value;
    }
    setLineItems(newItems);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    Array.from(files).forEach((file: File) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        setAttachments(prev => [...prev, {
          id: Math.random().toString(36).substr(2, 9),
          name: file.name,
          type: file.type,
          data: event.target?.result as string
        }]);
      };
      reader.readAsDataURL(file);
    });
  };

    const handleSubmit = async (e: React.FormEvent) => {
      e.preventDefault();
      if (!String(entity || '').trim()) {
        return toast.error('Please select an active entity before submitting this request.');
      }
      if (template.attachments_required && attachments.length === 0) {
        return toast.error('This workflow requires at least one attachment');
      }
      if ((isPR || isSR) && !signature) {
        return toast.error(isSR ? 'Signature is mandatory for stock requisitions' : 'Signature is mandatory for Purchase Requests');
      }
      if ((isPR || isPO || isSR) && !String(currency).trim()) {
        return toast.error('Please select a currency.');
      }
      if (isSR && !SECTION_OPTIONS.includes(String(section || "").trim().toUpperCase() as (typeof SECTION_OPTIONS)[number])) {
        return toast.error('Please select a section.');
      }
      if (needsApproverPicker && (assignedApproverId === '' || !Number.isFinite(Number(assignedApproverId)))) {
        return toast.error('Please select who should approve this request.');
      }
      if (isPR && !String(suggestedSupplier).trim()) {
        return toast.error('Please enter the suggested supplier for this PR.');
      }
      let taxRateNum = 0;
      if (isPR || isPO || isSR) {
        const trimmed = taxRateInput.trim();
        if (trimmed === '') {
          const lab = procurementTaxLabelForEntity(entity);
          return toast.error(`Please enter ${lab} rate as a percentage (e.g. 18 for 18%).`);
        }
        const taxPct = parseFloat(trimmed);
        if (Number.isNaN(taxPct) || taxPct < 0 || taxPct > 100) {
          const lab = procurementTaxLabelForEntity(entity);
          return toast.error(`${lab} rate must be between 0 and 100 (e.g. 18 for 18%).`);
        }
        taxRateNum = procurementPercentToUnitRate(taxPct);
      }
      let discountRateNum = 0;
      if (isPR || isPO || isSR) {
        const dtrim = discountRateInput.trim();
        if (dtrim !== '') {
          const discPct = parseFloat(dtrim);
          if (Number.isNaN(discPct) || discPct < 0 || discPct > 100) {
            return toast.error('Discount must be between 0 and 100 (e.g. 5 for 5%). Leave blank for no discount.');
          }
          discountRateNum = procurementPercentToUnitRate(discPct);
        }
      }
      const normalizedLineItems = normalizeProcurementCatalogLineItems(lineItems, costCenterOptions, isSR);
      setLoading(true);
      try {
        await api.request('/api/workflow-requests', {
          method: 'POST',
          body: JSON.stringify({ 
            template_id: template.id, 
            title, 
            details: isPR || isPO || isSR ? '' : details, 
            entity,
            currency: isPR || isPO || isSR ? String(currency).trim() : undefined,
            section: isSR ? sectionPayloadFromSelection(section) : undefined,
            line_items: normalizedLineItems,
            attachments,
            requester_signed: !!((isPR || isSR) && signature),
            requester_signature: (isPR || isSR) && signature ? signature : undefined,
            tax_rate: isPR || isPO || isSR ? taxRateNum : 0,
            discount_rate: isPR || isPO || isSR ? discountRateNum : 0,
            ...(needsApproverPicker && assignedApproverId !== ''
              ? { assigned_approver_id: Number(assignedApproverId) }
              : {}),
            ...(isPR ? { suggested_supplier: String(suggestedSupplier).trim() } : {}),
          }),
        });
        toast.success('Request submitted successfully!');
        onSuccess();
      } catch (err: any) {
        toast.error(err.message);
      } finally {
        setLoading(false);
      }
    };

  return (
    <div className="flex flex-col h-full min-h-0 bg-white rounded-2xl border border-zinc-200 shadow-sm overflow-hidden">
      <div className="shrink-0 px-5 sm:px-6 py-4 border-b border-zinc-100 bg-gradient-to-r from-zinc-50/80 to-white flex items-center gap-3">
        <div className="w-10 h-10 bg-indigo-100 rounded-lg flex items-center justify-center shrink-0">
          <Send className="w-5 h-5 text-indigo-600" />
        </div>
        <div className="min-w-0">
          <h2 className="text-lg sm:text-xl font-bold text-zinc-900 truncate">Start Request</h2>
          <p className="text-xs sm:text-sm text-zinc-500 truncate">Template: <span className="font-semibold text-indigo-600">{template.name}</span></p>
        </div>
      </div>

      <form onSubmit={handleSubmit} className="flex flex-col flex-1 min-h-0">
        <div className="flex-1 min-h-0 overflow-y-auto overscroll-contain">
          <div className="px-5 sm:px-6 py-5 space-y-5 max-w-[1600px] mx-auto">
            <div className="grid gap-5 lg:grid-cols-2">
              <div className="space-y-2 lg:col-span-2">
                <label className="text-sm font-medium text-zinc-700">Request Title</label>
                <input
                  type="text"
                  required
                  placeholder="e.g., Marketing Budget for Q3"
                  className="w-full px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                />
              </div>
              {!isPR && !isPO && !isSR && (
              <div className="space-y-2 lg:col-span-2">
                <label className="text-sm font-medium text-zinc-700">Details</label>
                <textarea
                  required
                  rows={3}
                  placeholder="Provide necessary information for the approvers..."
                  className="w-full px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500 resize-y min-h-[88px]"
                  value={details}
                  onChange={(e) => setDetails(e.target.value)}
                />
              </div>
              )}
            </div>

            {(isPR || isPO || isSR) && (
              <div className="space-y-4">
              <div className="grid sm:grid-cols-3 gap-4">
                <div className="space-y-2">
                  <label className="text-sm font-medium text-zinc-700">Currency</label>
                  <select
                    className="w-full px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                    value={currency}
                    onChange={(e) => setCurrency(e.target.value)}
                    required={isPR || isPO || isSR}
                  >
                    <option value="">Select currency…</option>
                    {CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-sm font-medium text-zinc-700">{procurementTaxRateFormLabel(entity)}</label>
                  <div className="flex flex-wrap items-center gap-3">
                    <input
                      type="text"
                      inputMode="decimal"
                      autoComplete="off"
                      placeholder="e.g. 18"
                      className="w-36 px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                      value={taxRateInput}
                      onChange={(e) => {
                        const v = e.target.value;
                        if (v === '' || /^\d*\.?\d*$/.test(v)) setTaxRateInput(v);
                      }}
                    />
                    <span className="text-sm text-zinc-500">
                      {(() => {
                        const t = taxRateInput.trim();
                        if (t === '') return 'Percent (18 = 18%)';
                        const n = parseFloat(t);
                        if (Number.isNaN(n)) return 'Percent (18 = 18%)';
                        return `${n}%`;
                      })()}
                    </span>
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-sm font-medium text-zinc-700">Discount (%)</label>
                  <div className="flex flex-wrap items-center gap-3">
                    <input
                      type="text"
                      inputMode="decimal"
                      autoComplete="off"
                      placeholder="e.g. 5 or leave blank"
                      className="w-36 px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                      value={discountRateInput}
                      onChange={(e) => {
                        const v = e.target.value;
                        if (v === '' || /^\d*\.?\d*$/.test(v)) setDiscountRateInput(v);
                      }}
                    />
                    <span className="text-sm text-zinc-500">
                      {(() => {
                        const t = discountRateInput.trim();
                        if (t === '') return 'Optional; blank = none';
                        const n = parseFloat(t);
                        if (Number.isNaN(n)) return 'Optional; blank = none';
                        return `${n}% off subtotal`;
                      })()}
                    </span>
                  </div>
                </div>
              </div>
              {isSR && (
                <div className="space-y-2 max-w-xl">
                  <label className="text-sm font-medium text-zinc-700">Section</label>
                  <select
                    className="w-full px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                    value={section}
                    onChange={(e) => setSection(e.target.value)}
                    required
                  >
                    <option value="" disabled>Select section...</option>
                    {SECTION_OPTIONS.map((opt) => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </div>
              )}
              {isPR && (
                <div className="space-y-2 max-w-2xl lg:col-span-2">
                  <label className="text-sm font-medium text-zinc-700">Suggested supplier</label>
                  <input
                    type="text"
                    required
                    placeholder="One supplier for this entire purchase requisition"
                    className="w-full px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                    value={suggestedSupplier}
                    onChange={(e) => setSuggestedSupplier(e.target.value)}
                  />
                  <p className="text-xs text-zinc-500">Applies to all line items on this PR.</p>
                </div>
              )}
              </div>
            )}

            {needsApproverPicker && (
              <div className="space-y-2 max-w-xl">
                <label className="text-sm font-medium text-zinc-700">Approver</label>
                <select
                  className="w-full px-4 py-2.5 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                  value={assignedApproverId === '' ? '' : String(assignedApproverId)}
                  onChange={(e) => {
                    const v = e.target.value;
                    setAssignedApproverId(v === '' ? '' : Number(v));
                  }}
                  required
                >
                  <option value="">Select approver…</option>
                  {eligibleApprovers.map((u) => (
                    <option key={u.id} value={u.id}>
                      {u.username}
                      {u.department ? ` — ${u.department}` : ''}
                    </option>
                  ))}
                </select>
                <p className="text-xs text-zinc-500">
                  You choose who acts at the approver step. Checker / SOM / director steps are unchanged.
                </p>
              </div>
            )}

            {(columns.length > 0) && (
              <div className="flex flex-col min-h-[min(50vh,520px)] -mx-5 sm:-mx-6 border-y border-zinc-200 bg-zinc-50/40">
                <div className="shrink-0 flex items-center justify-between gap-3 px-5 sm:px-6 py-3 border-b border-zinc-200 bg-white/90 backdrop-blur-sm sticky top-0 z-10">
                  <label className="text-sm font-semibold text-zinc-800">
                    {isPR ? 'PR line items' : isPO ? 'PO line items' : isSR ? 'Stock requisition line items' : 'Line items'}
                  </label>
                  <button
                    type="button"
                    onClick={addLineItem}
                    className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-lg font-semibold hover:bg-indigo-700 transition-colors shadow-sm"
                  >
                    + Add Row
                  </button>
                </div>
                <div className="flex-1 min-h-0 overflow-auto">
                  <table className="w-full text-sm border-collapse">
                    <thead className="sticky top-0 z-[1] bg-zinc-100 shadow-[0_1px_0_0_rgb(228_228_231)]">
                      <tr>
                        <th className="px-3 sm:px-4 py-3 font-bold text-zinc-700 w-11 text-center border-b border-zinc-200">No</th>
                        {columns.map(col => (
                          <th key={col} className="px-2 sm:px-3 py-3 font-bold text-zinc-700 text-left min-w-[140px] border-b border-zinc-200">{col}</th>
                        ))}
                        <th className="px-3 py-3 font-bold text-zinc-700 w-14 text-center border-b border-zinc-200">Del</th>
                      </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-zinc-100">
                      {lineItems.length === 0 && (
                        <tr>
                          <td colSpan={columns.length + 2} className="py-16 text-center align-middle">
                            <Package className="w-12 h-12 text-zinc-300 mx-auto mb-3" />
                            <p className="text-zinc-500">No items yet. Click &quot;+ Add Row&quot; to start.</p>
                          </td>
                        </tr>
                      )}
                      {lineItems.map((item, idx) => (
                        <React.Fragment key={item.id || idx}>
                        <tr className="hover:bg-indigo-50/40 transition-colors">
                          <td className="px-3 sm:px-4 py-2.5 text-center text-zinc-500 font-medium align-top">{idx + 1}</td>
                          {columns.map(col => (
                            <td key={col} className="px-2 py-2 align-top">
                              {isCostCenterGridColumn(col) ? (
                                <SearchableSelect
                                  options={costCenterOptions}
                                  value={item[col] ?? ''}
                                  onChange={(val) => updateLineItem(idx, col, val)}
                                  placeholder="Select Cost Center..."
                                  loading={costCentersLoading}
                                  loadFailed={costCentersLoadFailed}
                                  loadErrorHint={costCentersFetchError}
                                />
                              ) : (isSR && isSpareLocationColumn(col)) ? (
                                <SearchableSelect
                                  options={costCenterOptions}
                                  value={item[col] ?? ''}
                                  onChange={(val) => updateLineItem(idx, col, val)}
                                  placeholder="Select Spare for (Location)..."
                                  loading={costCentersLoading}
                                  loadFailed={costCentersLoadFailed}
                                  loadErrorHint={costCentersFetchError}
                                />
                              ) : (
                                <input
                                  type={col === 'Delivery Date' || col === 'Request to be delivered on' ? 'date' : (isProcurementNumericGridColumn(col) ? 'number' : 'text')}
                                  min={isProcurementQuantityGridColumn(col) ? '0' : undefined}
                                  step={col.trim().toLowerCase() === 'quantity' || col.trim().toLowerCase() === 'min quantity' || col.trim().toLowerCase() === 'max quantity' ? '1' : (col.trim().toLowerCase() === 'unit price' || col.trim().toLowerCase() === 'price' || col.trim().toLowerCase() === 'amount' ? '0.01' : undefined)}
                                  className="w-full min-w-0 px-2.5 py-2 rounded-lg border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
                                  value={col === REMARKS_LINE_COL ? lineItemRemarksDisplay(item) : (item[col] ?? '')}
                                  onChange={(e) => updateLineItem(idx, col, e.target.value)}
                                  placeholder={col}
                                />
                              )}
                            </td>
                          ))}
                          <td className="px-2 py-2 text-center align-top">
                            <button 
                              type="button" 
                              onClick={() => removeLineItem(idx)} 
                              className="p-2 text-zinc-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all"
                              title="Remove row"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                        <tr className="bg-zinc-50/80 border-b border-zinc-100">
                          <td colSpan={columns.length + 2} className="px-4 sm:px-5 py-2 align-top">
                            <label className="text-[10px] font-semibold text-zinc-500 uppercase tracking-wide">Line remarks <span className="font-normal normal-case text-zinc-400">(optional, multiple lines)</span></label>
                            <textarea
                              rows={3}
                              className="w-full mt-1.5 px-3 py-2 rounded-lg border border-zinc-200 text-sm text-zinc-800 outline-none focus:ring-2 focus:ring-indigo-500 bg-white resize-y min-h-[72px] placeholder:text-zinc-400"
                              value={item[LINE_ITEM_REMARKS_KEY] ?? ''}
                              onChange={(e) => updateLineItem(idx, LINE_ITEM_REMARKS_KEY, e.target.value)}
                              placeholder="Extra notes for this line only…"
                            />
                          </td>
                        </tr>
                        </React.Fragment>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="space-y-2">
              <label className="text-sm font-medium text-zinc-700">Attachments</label>
              <div className="flex flex-wrap gap-2 mb-2">
                {attachments.map((file) => (
                  <div key={file.id} className="flex items-center gap-2 bg-zinc-100 px-3 py-1 rounded-full text-xs text-zinc-600 border border-zinc-200">
                    <Paperclip className="w-3 h-3" />
                    {file.name}
                    <button type="button" onClick={() => setAttachments(attachments.filter((a) => a.id !== file.id))}>
                      <Trash2 className="w-3 h-3 hover:text-red-500" />
                    </button>
                  </div>
                ))}
              </div>
              <label className="flex flex-col items-center justify-center w-full h-24 border-2 border-dashed border-zinc-300 rounded-xl cursor-pointer hover:bg-zinc-50 transition-colors">
                <div className="flex flex-col items-center justify-center">
                  <Upload className="w-6 h-6 text-zinc-400 mb-1" />
                  <p className="text-xs text-zinc-500">Upload supporting documents</p>
                </div>
                <input type="file" className="hidden" multiple onChange={handleFileChange} />
              </label>
            </div>

            {(isPR || isSR) && (
              <div className="space-y-2">
                <label className="text-sm font-medium text-zinc-700 flex items-center gap-2">
                  Signature (Preparer)
                  <span className="text-[10px] bg-red-100 text-red-600 px-1.5 py-0.5 rounded uppercase font-bold">Mandatory</span>
                </label>
                <SignaturePad
                  onSave={setSignature}
                  onClear={() => setSignature(null)}
                  value={signature}
                  savedSignature={savedSignature}
                  onUseSaved={useSavedSignature}
                  onSaveDefault={savedSignatureLoading ? undefined : saveSignatureAsDefault}
                  onClearSaved={savedSignatureLoading ? undefined : removeSavedSignature}
                />
                {savedSignatureLoading ? (
                  <p className="text-[11px] text-zinc-400">Syncing saved signature…</p>
                ) : null}
              </div>
            )}

            <div className="bg-zinc-50 p-4 rounded-xl border border-zinc-200">
              <h4 className="text-xs font-bold text-zinc-400 uppercase tracking-wider mb-3">Approval Steps for this Request</h4>
              <div className="space-y-2">
                {template.steps.map((step, i) => (
                  <div key={step.id} className="flex items-center gap-3 text-sm">
                    <div className="w-5 h-5 rounded-full bg-zinc-200 text-zinc-600 flex items-center justify-center text-[10px] font-bold shrink-0">
                      {i + 1}
                    </div>
                    <span className="text-zinc-700">{step.label}</span>
                    <span className="text-zinc-400 text-xs">({step.approverRole})</span>
                  </div>
                ))}
                {isPO && (
                  <p className="text-[11px] text-zinc-500 mt-2">
                    The Director step is included only when the order total (including tax), converted to MYR, exceeds RM 30,000.
                  </p>
                )}
              </div>
            </div>
          </div>
        </div>

        <div className="shrink-0 border-t border-zinc-200 bg-zinc-50/90 px-5 sm:px-6 py-4">
          <button
            type="submit"
            disabled={loading}
            className="w-full bg-indigo-600 text-white py-3.5 rounded-xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 disabled:opacity-50"
          >
            {loading ? 'Submitting...' : 'Submit Request'}
          </button>
        </div>
      </form>
    </div>
  );
};

const WorkflowRequestList = ({ 
  requests, 
  user, 
  onRefresh,
  preSelectedRequestId,
  onClearPreSelected
}: { 
  requests: WorkflowRequest[], 
  user: User, 
  onRefresh: () => void,
  preSelectedRequestId?: number | null,
  onClearPreSelected?: () => void
}) => {
  const [selectedRequest, setSelectedRequest] = useState<WorkflowRequest | null>(null);
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [approvals, setApprovals] = useState<RequestApproval[]>([]);
  const [comment, setComment] = useState('');
  const [approverSignature, setApproverSignature] = useState<string | null>(null);
  const [savedApproverSignature, setSavedApproverSignature] = useState<string | null>(null);
  const [savedApproverSignatureLoading, setSavedApproverSignatureLoading] = useState(false);
  const [loading, setLoading] = useState(false);
  const {
    options: costCenterOptions,
    loading: costCentersLoading,
    failed: costCentersLoadFailed,
    fetchError: costCentersFetchError,
  } = useCostCenterOptions();
  const [filterStatus, setFilterStatus] = useState<string>('all');
  const [filterDept, setFilterDept] = useState<string>('all');
  const [filterEntity, setFilterEntity] = useState<string>('all');
  const [filterChosenApprover, setFilterChosenApprover] = useState<string>('all');
  const [listSearch, setListSearch] = useState('');
  const [sortBy, setSortBy] = useState<
    'request' | 'entity' | 'department' | 'requester' | 'chosenApprover' | 'status' | 'date'
  >('date');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('desc');
  const [isEditing, setIsEditing] = useState(false);
  const [editData, setEditData] = useState({
    title: '',
    details: '',
    line_items: [] as any[],
    currency: '',
    cost_center: '',
    section: '',
    suggested_supplier: '',
    tax_rate: 18,
    discount_rate: 0,
    requester_username: '',
    requester_name: '',
  });
  const [editAttachmentKeep, setEditAttachmentKeep] = useState<Attachment[]>([]);
  const [editAttachmentAdd, setEditAttachmentAdd] = useState<{ id: string; name: string; type: string; data: string }[]>([]);
  const [viewingPdf, setViewingPdf] = useState<{ url: string; fileName: string } | null>(null);
  const [detailsViewMode, setDetailsViewMode] = useState<'details' | 'pdf'>('details');
  const [detailsPdfPreview, setDetailsPdfPreview] = useState<{ url: string; fileName: string } | null>(null);
  const [convertPoModal, setConvertPoModal] = useState<ConvertPoModalTarget | null>(null);
  const [onBehalfOptions, setOnBehalfOptions] = useState<{ id: number; username: string; department: string }[]>([]);
  const [onBehalfApproverId, setOnBehalfApproverId] = useState<number | ''>('');
  const [onBehalfLoading, setOnBehalfLoading] = useState(false);

  useEffect(() => {
    if (preSelectedRequestId) {
      const request = requests.find(r => r.id === preSelectedRequestId);
      if (request) {
        handleViewRequest(request);
        if (onClearPreSelected) onClearPreSelected();
      }
    }
  }, [preSelectedRequestId, requests]);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setSavedApproverSignatureLoading(true);
        const r = await api.request("/api/me/signature", { skipEntity: true });
        if (cancelled) return;
        setSavedApproverSignature(r?.exists && typeof r?.dataUrl === "string" ? r.dataUrl : null);
      } catch {
        if (!cancelled) setSavedApproverSignature(null);
      } finally {
        if (!cancelled) setSavedApproverSignatureLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const prAtApproverStepForBehalf =
    !!selectedRequest &&
    isPRRequest(selectedRequest) &&
    isWorkflowRequestPending(selectedRequest) &&
    (() => {
      const st = selectedRequest.template_steps[selectedRequest.current_step_index];
      return st?.approverRole?.toLowerCase() === 'approver';
    })();
  const mustPickOnBehalf = prAtApproverStepForBehalf && isPrSignOnBehalfUser(user);
  const assignedApproverForSelected =
    mustPickOnBehalf && selectedRequest?.assigned_approver_id != null
      ? Number(selectedRequest.assigned_approver_id)
      : NaN;
  const hasFixedAssignedApprover =
    mustPickOnBehalf && Number.isFinite(assignedApproverForSelected) && assignedApproverForSelected > 0;

  useEffect(() => {
    if (!mustPickOnBehalf || !selectedRequest) {
      setOnBehalfOptions([]);
      setOnBehalfApproverId('');
      return;
    }
    let cancelled = false;
    (async () => {
      setOnBehalfLoading(true);
      try {
        api.setActiveEntity(String(selectedRequest.entity || '').trim() || null);
        const list = await api.request(`/api/workflow-requests/${selectedRequest.id}/on-behalf-approver-options`);
        if (!cancelled && Array.isArray(list)) setOnBehalfOptions(list);
      } catch {
        if (!cancelled) setOnBehalfOptions([]);
      } finally {
        if (!cancelled) setOnBehalfLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [mustPickOnBehalf, selectedRequest?.id]);

  useEffect(() => {
    if (!mustPickOnBehalf) return;
    if (hasFixedAssignedApprover) {
      setOnBehalfApproverId(assignedApproverForSelected);
    }
  }, [mustPickOnBehalf, hasFixedAssignedApprover, assignedApproverForSelected]);

  const hasPermission = (permission: string) => {
    if (!user) return false;
    if (user.permissions?.includes('admin')) return true;
    return user.permissions?.includes(permission);
  };

  const isAdmin = user.roles?.some(r => r.toLowerCase() === 'admin') || hasPermission('admin');
  const isPurchasing = user.roles?.some(r => r.toLowerCase() === 'purchasing');
  const isDirector = user.roles?.some(r => r.toLowerCase() === 'director') && (user.department || '').toLowerCase() === 'management';

  const departments = Array.from(new Set(requests.map(r => r.department)));
  const entities = Array.from(
    new Set(
      requests
        .map((r) => String(r.entity || '').trim())
        .filter((e) => !!e)
    )
  );
  const chosenApproverOptions = Array.from(
    new Map(
      requests
        .filter((r) => requestHasChosenApprover(r))
        .map((r) => [Number(r.assigned_approver_id), chosenApproverNameLabel(r)] as [number, string])
    ).entries()
  )
    .map(([id, label]) => ({ id, label }))
    .sort((a, b) => a.label.localeCompare(b.label, undefined, { sensitivity: 'base' }));

  const searchQ = listSearch.trim().toLowerCase();
  const filteredRequests = requests.filter(r => {
    const statusMatch = filterStatus === 'all' || normalizeWorkflowRequestStatus(r.status) === filterStatus;
    const deptMatch = filterDept === 'all' || r.department === filterDept;
    const entityMatch = filterEntity === 'all' || String(r.entity || '').trim() === filterEntity;
    const chosenApproverMatch =
      filterChosenApprover === 'all' ||
      (filterChosenApprover === 'none' && !requestHasChosenApprover(r)) ||
      (filterChosenApprover !== 'none' &&
        requestHasChosenApprover(r) &&
        String(r.assigned_approver_id) === filterChosenApprover);
    const searchMatch =
      !searchQ ||
      [
        r.title,
        displayRequestSerial(r),
        String(r.formatted_id ?? ''),
        r.template_name,
        r.requester_name,
        chosenApproverNameLabel(r),
        String(r.entity ?? ''),
        r.department,
        String(r.assigned_approver_designation ?? ''),
      ]
        .filter(Boolean)
        .join(' ')
        .toLowerCase()
        .includes(searchQ);
    return statusMatch && deptMatch && entityMatch && chosenApproverMatch && searchMatch;
  });

  const toggleSort = (
    next: 'request' | 'entity' | 'department' | 'requester' | 'chosenApprover' | 'status' | 'date'
  ) => {
    if (sortBy === next) {
      setSortDirection((d) => (d === 'asc' ? 'desc' : 'asc'));
      return;
    }
    setSortBy(next);
    setSortDirection(next === 'date' ? 'desc' : 'asc');
  };

  const sortedFilteredRequests = [...filteredRequests].sort((a, b) => {
    const dir = sortDirection === 'asc' ? 1 : -1;
    const byText = (x: string, y: string) => x.localeCompare(y, undefined, { sensitivity: 'base' }) * dir;
    if (sortBy === 'date') {
      return (new Date(a.created_at).getTime() - new Date(b.created_at).getTime()) * dir;
    }
    if (sortBy === 'request') {
      const at = `${displayRequestSerial(a)} ${a.title || ''}`.trim();
      const bt = `${displayRequestSerial(b)} ${b.title || ''}`.trim();
      return byText(at, bt);
    }
    if (sortBy === 'entity') return byText(String(a.entity || ''), String(b.entity || ''));
    if (sortBy === 'department') return byText(String(a.department || ''), String(b.department || ''));
    if (sortBy === 'requester') return byText(String(a.requester_name || ''), String(b.requester_name || ''));
    if (sortBy === 'chosenApprover') {
      return byText(chosenApproverNameLabel(a), chosenApproverNameLabel(b));
    }
    return byText(formatWorkflowRequestStatusLabel(a), formatWorkflowRequestStatusLabel(b));
  });

  const sortIndicator = (
    key: 'request' | 'entity' | 'department' | 'requester' | 'chosenApprover' | 'status' | 'date'
  ) =>
    sortBy === key ? (sortDirection === 'asc' ? ' ↑' : ' ↓') : '';

  const handleViewRequest = async (request: WorkflowRequest) => {
    api.setActiveEntity(String(request.entity || '').trim() || null);
    setSelectedRequest(request);
    setEditData({
      title: request.title,
      details: request.details,
      line_items: request.line_items || [],
      currency: request.currency?.trim() ?? '',
      cost_center: request.cost_center || '',
      section: sectionSelectionFromStored(request.section),
      tax_rate: procurementUnitRateToPercent(request.tax_rate, 0.18),
      discount_rate: procurementUnitRateToPercent(request.discount_rate, 0),
      requester_username: String(request.requester_username ?? '').trim(),
      requester_name: String(request.requester_name ?? '').trim(),
    });
    setComment('');
    setApproverSignature(null);
    setOnBehalfApproverId('');
    setIsEditing(false);
    setDetailsViewMode('details');
    setDetailsPdfPreview((prev) => {
      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
      return null;
    });
    await fetchDetails(request.id);
  };

  const handlePrintPR = (request: WorkflowRequest) => {
    const merged = mergeRequestApprovalsForPdf(request, selectedRequest?.id, approvals);
    const preview = createWorkflowRequestPdfPreview(merged);
    void persistGeneratedProcurementFormPdf(request.id, preview.pdfDataUrl);
    setViewingPdf((prev) => {
      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
      return { url: preview.url, fileName: preview.fileName };
    });
    toast.success('PDF ready to view');
  };

  const handleShowRequestPdfInline = (request: WorkflowRequest) => {
    const merged = mergeRequestApprovalsForPdf(request, selectedRequest?.id, approvals);
    const preview = createWorkflowRequestPdfPreview(merged);
    void persistGeneratedProcurementFormPdf(request.id, preview.pdfDataUrl);
    setDetailsPdfPreview((prev) => {
      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
      return { url: preview.url, fileName: preview.fileName };
    });
    setDetailsViewMode('pdf');
  };

  const fetchDetails = async (id: number) => {
    try {
      const [attData, appData] = await Promise.all([
        api.request(`/api/workflow-requests/${id}/attachments`),
        api.request(`/api/workflow-requests/${id}/approvals`)
      ]);
      setAttachments(attData);
      setApprovals(appData);
    } catch (err) {
      toast.error('Failed to load request details');
    }
  };

  const handleEditAttachmentAdd = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;
    Array.from(files).forEach((file: File) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        setEditAttachmentAdd((prev) => [
          ...prev,
          {
            id: Math.random().toString(36).substr(2, 9),
            name: file.name,
            type: file.type,
            data: event.target?.result as string,
          },
        ]);
      };
      reader.readAsDataURL(file);
    });
    e.target.value = '';
  };

  const handleSaveEdit = async () => {
    if (!selectedRequest) return;
    if (
      (isPRRequest(selectedRequest) || isPO_Only(selectedRequest) || isSRRequest(selectedRequest)) &&
      !String(editData.currency).trim()
    ) {
      return toast.error('Please select a currency.');
    }
    if (
      isSRRequest(selectedRequest) &&
      !SECTION_OPTIONS.includes(String(editData.section || "").trim().toUpperCase() as (typeof SECTION_OPTIONS)[number])
    ) {
      return toast.error('Please select a section.');
    }
    setLoading(true);
    try {
      const detailsOut = isProcurementPRorPORequest(selectedRequest) ? '' : editData.details;
      const taxUnit = procurementPercentToUnitRate(Number(editData.tax_rate) || 0);
      const discUnit = procurementPercentToUnitRate(Number(editData.discount_rate) || 0);
      const sectionOut = isSRRequest(selectedRequest) ? sectionPayloadFromSelection(editData.section) : editData.section;
      const normalizedLineItems = normalizeProcurementCatalogLineItems(
        editData.line_items,
        costCenterOptions,
        isSRRequest(selectedRequest)
      );
      const curOut =
        isPRRequest(selectedRequest) || isPO_Only(selectedRequest) || isSRRequest(selectedRequest)
          ? String(editData.currency).trim()
          : undefined;
      if (isPRRequest(selectedRequest) && !String(editData.suggested_supplier ?? '').trim()) {
        return toast.error('Please enter the suggested supplier for this PR.');
      }
      if (selectedRequest.attachments_required && (editAttachmentKeep.length + editAttachmentAdd.length) === 0) {
        return toast.error('This workflow requires at least one attachment.');
      }
      await api.request(`/api/workflow-requests/${selectedRequest.id}`, {
        method: 'PATCH',
        body: JSON.stringify({
          ...editData,
          line_items: normalizedLineItems,
          tax_rate: taxUnit,
          discount_rate: discUnit,
          details: detailsOut,
          currency: curOut,
          section: sectionOut,
          attachment_keep_ids: editAttachmentKeep.map((a) => a.id),
          attachments_add: editAttachmentAdd.map((a) => ({ name: a.name, type: a.type, data: a.data })),
          ...(isPRRequest(selectedRequest)
            ? { suggested_supplier: String(editData.suggested_supplier ?? '').trim() }
            : {}),
        }),
      });
      toast.success('Request updated');
      setIsEditing(false);
      setEditAttachmentAdd([]);
      onRefresh();
      // Update local selected request
      setSelectedRequest({
        ...selectedRequest,
        ...editData,
        line_items: normalizedLineItems,
        tax_rate: taxUnit,
        discount_rate: discUnit,
        currency: curOut !== undefined ? curOut : selectedRequest.currency,
        details: detailsOut,
        section: sectionOut,
        ...(isPRRequest(selectedRequest)
          ? { suggested_supplier: String(editData.suggested_supplier ?? '').trim() }
          : {}),
      });
      await fetchDetails(selectedRequest.id);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleApprove = async (id: number, status: 'approved' | 'rejected') => {
    const needSig = selectedRequest && requiresProcurementApproverSignaturePad(selectedRequest);
    if (status === 'approved' && needSig && !approverSignature) {
      return toast.error('Signature is required to approve this document');
    }
    if (
      mustPickOnBehalf &&
      (onBehalfApproverId === '' || !Number.isFinite(Number(onBehalfApproverId)))
    ) {
      return toast.error('Select which approver you are signing on behalf of');
    }
    setLoading(true);
    try {
      const sendSig = needSig && approverSignature ? approverSignature : undefined;
      const approveResult = await api.request(`/api/workflow-requests/${id}/approve`, {
        method: 'POST',
        body: JSON.stringify({
          status,
          comment,
          approver_signed: !!sendSig,
          approver_signature: sendSig,
          ...(mustPickOnBehalf && onBehalfApproverId !== ''
            ? { on_behalf_of_approver_id: Number(onBehalfApproverId) }
            : {}),
        }),
      });
      if (status === 'approved' && selectedRequest && approveResult?.reached_final_approval === true) {
        await persistPrSrFormPdfAfterWorkflowCompleted(id, selectedRequest, { skipFinalStepCheck: true });
      }
      toast.success(`Request ${status}`);
      onRefresh();
      setSelectedRequest(null);
      setComment('');
      setApproverSignature(null);
      setOnBehalfApproverId('');
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const useSavedApproverSignature = async () => {
    try {
      setSavedApproverSignatureLoading(true);
      const r = await api.request("/api/me/signature", { skipEntity: true });
      const next = r?.exists && typeof r?.dataUrl === "string" ? r.dataUrl : null;
      setSavedApproverSignature(next);
      if (!next) {
        toast.error("No saved signature found.");
        return;
      }
      setApproverSignature(next);
      toast.success("Using saved signature.");
    } catch (e: any) {
      toast.error(e?.message || "Failed to load saved signature");
    } finally {
      setSavedApproverSignatureLoading(false);
    }
  };

  const saveApproverSignatureAsDefault = async () => {
    if (!approverSignature || !isSignatureImageDataUrl(approverSignature)) {
      toast.error("Please draw or upload a signature first.");
      return;
    }
    try {
      setSavedApproverSignatureLoading(true);
      await api.request("/api/me/signature", {
        method: "PUT",
        skipEntity: true,
        body: JSON.stringify({ dataUrl: approverSignature }),
      });
      setSavedApproverSignature(approverSignature);
      toast.success("Saved signature.");
    } catch (e: any) {
      toast.error(e?.message || "Failed to save signature");
    } finally {
      setSavedApproverSignatureLoading(false);
    }
  };

  const removeSavedApproverSignature = async () => {
    try {
      setSavedApproverSignatureLoading(true);
      await api.request("/api/me/signature", { method: "DELETE", skipEntity: true });
      setSavedApproverSignature(null);
      toast.success("Removed saved signature.");
    } catch (e: any) {
      toast.error(e?.message || "Failed to remove saved signature");
    } finally {
      setSavedApproverSignatureLoading(false);
    }
  };

  const handleResubmit = async (id: number) => {
    setLoading(true);
    try {
      await api.request(`/api/workflow-requests/${id}/resubmit`, {
        method: 'POST',
      });
      toast.success('Request resubmitted successfully!');
      onRefresh();
      setSelectedRequest(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleRequesterCancel = async (id: number) => {
    const promptResult = window.prompt('Optional - Cancellation Reason:', '');
    if (promptResult === null) return;
    const reason = promptResult.trim();
    const confirmed = window.confirm('Cancel this request? This action is final and cannot be resubmitted.');
    if (!confirmed) return;
    setLoading(true);
    try {
      await api.request(`/api/workflow-requests/${id}/cancel`, {
        method: 'POST',
        body: JSON.stringify({ comment: reason }),
      });
      toast.success('Request cancelled');
      onRefresh();
      setSelectedRequest(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleGeneratePODraft = (request: WorkflowRequest) => {
    const doc = new jsPDF();
    const margin = 20;
    let y = 20;

    // Header
    doc.setFontSize(20);
    doc.setTextColor(79, 70, 229); // Indigo-600
    doc.text('PURCHASE ORDER DRAFT', margin, y);
    y += 10;

    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Generated on: ${formatDateTimeMYT(new Date())}`, margin, y);
    y += 15;

    // Request Info
    doc.setFontSize(12);
    doc.setTextColor(0);
    doc.setFont('helvetica', 'bold');
    doc.text('Request Information', margin, y);
    y += 7;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.text(`PR ID: ${displayRequestSerial(request)}`, margin, y);
    y += 5;
    doc.text(`Title: ${request.title}`, margin, y);
    y += 5;
    doc.text(`Department: ${request.department}`, margin, y);
    y += 5;
    doc.text(`Entity: ${entityLegalDisplayName(request.entity)} (${request.entity?.trim() || '-'})`, margin, y);
    y += 15;

    // Line Items Table
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text('Order Items', margin, y);
    y += 7;

    const headers = ['Item', 'Qty', 'Unit Price', 'Total'];
    const colWidths = [80, 20, 35, 35];
    
    // Table Header
    doc.setFillColor(244, 244, 245);
    doc.rect(margin, y - 5, 170, 7, 'F');
    doc.setFontSize(9);
    let x = margin;
    headers.forEach((h, i) => {
      doc.text(h, x + 2, y);
      x += colWidths[i];
    });
    y += 7;

    // Table Rows
    doc.setFont('helvetica', 'normal');
    request.line_items?.forEach((item) => {
      const qty = parseFloat(item['Quantity'] || '0');
      const price = parseFloat(item['Unit Price'] || '0');
      const rowTot = qty * price;

      x = margin;
      doc.text(String(item['Item'] || '-'), x + 2, y, { maxWidth: colWidths[0] - 4 });
      doc.text(String(qty), x + colWidths[0] + 2, y);
      doc.text(`${procurementCurrencyPrefix(request.currency)}${formatProcurementMoney(price)}`.trim(), x + colWidths[0] + colWidths[1] + 2, y);
      doc.text(`${procurementCurrencyPrefix(request.currency)}${formatProcurementMoney(rowTot)}`.trim(), x + colWidths[0] + colWidths[1] + colWidths[2] + 2, y);
      y += 7;
    });

    y += 5;
    const m = procurementMoneyTotals(request);
    const curp = procurementCurrencyPrefix(request.currency);
    doc.setFont('helvetica', 'bold');
    doc.text(`Subtotal: ${curp}${formatProcurementMoney(m.subtotal)}`.trim(), 110, y);
    y += 5;
    doc.text(`Discount (${(m.discountRate * 100).toFixed(0)}%): ${curp}${formatProcurementMoney(m.discountAmount)}`.trim(), 110, y);
    y += 5;
    doc.text(`${procurementTaxLabelForEntity(request.entity)} (${(m.taxRate * 100).toFixed(0)}%): ${curp}${formatProcurementMoney(m.taxAmount)}`.trim(), 110, y);
    y += 7;
    doc.setFontSize(12);
    doc.text(`TOTAL AMOUNT: ${curp}${formatProcurementMoney(m.total)}`.trim(), 110, y);

    doc.save(`PO_Draft_${displayRequestSerial(request)}.pdf`);
    toast.success('PO Draft PDF generated');
  };

  const handleUploadRealPoAttachment = async (file: File) => {
    if (!selectedRequest || !isPO_Only(selectedRequest)) return;
    setLoading(true);
    try {
      const data = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          if (typeof reader.result === 'string') resolve(reader.result);
          else reject(new Error('Failed to read selected file'));
        };
        reader.onerror = () => reject(new Error('Failed to read selected file'));
        reader.readAsDataURL(file);
      });
      await api.request(`/api/workflow-requests/${selectedRequest.id}`, {
        method: 'PATCH',
        body: JSON.stringify({
          attachments_add: [{ name: file.name, type: file.type || 'application/octet-stream', data }],
        }),
      });
      toast.success('PO document uploaded');
      await fetchDetails(selectedRequest.id);
      onRefresh();
    } catch (err: any) {
      toast.error(err?.message || 'Failed to upload PO document');
    } finally {
      setLoading(false);
    }
  };

  const openConvertPoModal = (r: WorkflowRequest) => {
    setConvertPoModal({
      id: r.id,
      title: r.title || '',
      prSerial: r.formatted_id ? toUpperSerial(r.formatted_id) : `#${r.id}`,
      entityCode: String(r.entity || '').trim() || '—',
    });
  };

  const handleConvertToPOConfirm = async (poNumber: string, poUpload?: ConvertPoUploadPayload) => {
    if (!convertPoModal) return;
    setLoading(true);
    try {
      const result = await api.request(`/api/workflow-requests/${convertPoModal.id}/convert-to-po`, {
        method: 'POST',
        body: JSON.stringify({ po_number: poNumber, po_upload: poUpload || null }),
      });
      toast.success(
        result?.merged_into_existing
          ? `PR appended to PO: ${toUpperSerial(result.formatted_id)}`
          : `PO created: ${toUpperSerial(result.formatted_id)}`
      );
      setConvertPoModal(null);
      onRefresh();
      setSelectedRequest(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handlePurchasingFinalDecision = async (id: number, decision: 'rejected' | 'cancelled') => {
    const decisionComment =
      window.prompt(
        decision === 'cancelled'
          ? 'Cancellation reason (optional, recommended for audit):'
          : 'Rejection reason (optional):',
        ''
      ) ?? '';
    if (decision === 'cancelled' && !String(decisionComment || '').trim()) {
      return toast.error('Please enter a cancellation reason.');
    }
    setLoading(true);
    try {
      await api.request(`/api/workflow-requests/${id}/purchasing-decision`, {
        method: 'POST',
        body: JSON.stringify({
          decision,
          comment: decisionComment,
        }),
      });
      toast.success(decision === 'cancelled' ? 'PR cancelled' : 'PR rejected by purchasing');
      onRefresh();
      setSelectedRequest(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const currentStep = selectedRequest ? selectedRequest.template_steps[selectedRequest.current_step_index] : null;
  const canApprove =
    selectedRequest &&
    currentStep &&
    userCanApproveWorkflowStep(user, selectedRequest, currentStep);

  const canEdit = selectedRequest && (
    (isWorkflowRequestPending(selectedRequest) && (
      user.id === selectedRequest.requester_id ||
      (canApprove && hasPermission('edit_requests')) ||
      (isPurchasing && isPO_Only(selectedRequest))
    )) ||
    // Requester may edit rejected PRs before resubmitting.
    (isWorkflowRequestRejected(selectedRequest) && isPR_Only(selectedRequest) && user.id === selectedRequest.requester_id) ||
    (isWorkflowRequestFullyApproved(selectedRequest) && isPR_Only(selectedRequest) && isPurchasing)
  );
  const canRequesterCancelSelected =
    !!selectedRequest && canRequesterCancelPendingRequest(selectedRequest, user, approvals);

  return (
    <div className="space-y-3 flex flex-col flex-1 min-h-0">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 shrink-0">
        <h2 className="text-xl font-bold text-zinc-900">
          {(isAdmin || isDirector || hasPermission('view_history')) ? 'All Requests' : 'My Requests'}
        </h2>
        
        <div className="flex flex-col gap-2 w-full md:w-auto md:items-end">
          <div className="flex items-center gap-2 bg-white border border-zinc-200 rounded-lg px-3 py-1.5 w-full md:w-72 md:max-w-sm">
            <Search className="w-3.5 h-3.5 text-zinc-400 shrink-0" />
            <input
              type="search"
              enterKeyHint="search"
              placeholder="Search by ref, title, requester, entity…"
              className="text-xs bg-transparent border-none focus:ring-0 outline-none text-zinc-700 font-medium w-full placeholder:text-zinc-400"
              value={listSearch}
              onChange={(e) => setListSearch(e.target.value)}
              aria-label="Search requests"
            />
          </div>
          <div className="flex flex-wrap items-center gap-2">
          <div className="flex items-center gap-2 bg-white border border-zinc-200 rounded-lg px-3 py-1.5">
            <Filter className="w-3 h-3 text-zinc-400" />
            <select 
              className="text-xs bg-transparent border-none focus:ring-0 outline-none text-zinc-600 font-medium"
              value={filterStatus}
              onChange={(e) => setFilterStatus(e.target.value)}
            >
              <option value="all">All Status</option>
              <option value="pending">Pending</option>
              <option value="approved">Approved</option>
              <option value="rejected">Rejected</option>
              <option value="cancelled">Cancelled</option>
            </select>
          </div>

          <div className="flex items-center gap-2 bg-white border border-zinc-200 rounded-lg px-3 py-1.5">
            <Building2 className="w-3 h-3 text-zinc-400" />
            <select 
              className="text-xs bg-transparent border-none focus:ring-0 outline-none text-zinc-600 font-medium"
              value={filterDept}
              onChange={(e) => setFilterDept(e.target.value)}
            >
              <option value="all">All Depts</option>
              {departments.map(dept => (
                <option key={dept} value={dept}>{dept}</option>
              ))}
            </select>
          </div>
          <div className="flex items-center gap-2 bg-white border border-zinc-200 rounded-lg px-3 py-1.5">
            <Shield className="w-3 h-3 text-zinc-400" />
            <select
              className="text-xs bg-transparent border-none focus:ring-0 outline-none text-zinc-600 font-medium"
              value={filterEntity}
              onChange={(e) => setFilterEntity(e.target.value)}
            >
              <option value="all">All Entities</option>
              {entities.map((ent) => (
                <option key={ent} value={ent}>{ent}</option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-2 bg-white border border-zinc-200 rounded-lg px-3 py-1.5 min-w-0">
            <UserPlus className="w-3 h-3 text-zinc-400 shrink-0" />
            <select
              className="text-xs bg-transparent border-none focus:ring-0 outline-none text-zinc-600 font-medium max-w-[11rem] sm:max-w-[14rem] truncate"
              value={filterChosenApprover}
              onChange={(e) => setFilterChosenApprover(e.target.value)}
              title="Filter by chosen approver"
            >
              <option value="all">All chosen approvers</option>
              <option value="none">No chosen approver</option>
              {chosenApproverOptions.map((o) => (
                <option key={o.id} value={String(o.id)}>{o.label}</option>
              ))}
            </select>
          </div>

          <span className="text-[10px] text-zinc-400 font-bold uppercase ml-2">
              {sortedFilteredRequests.length} Results
          </span>

          <button
            type="button"
            onClick={() => onRefresh()}
            title="Reload list from server"
            className="flex items-center gap-1 text-[10px] font-bold text-zinc-600 hover:text-indigo-600 bg-zinc-100 hover:bg-indigo-50 border border-zinc-200 px-2 py-1 rounded-lg transition-colors"
          >
            <RefreshCw className="w-3 h-3 shrink-0" />
            Reload
          </button>

          {(filterStatus !== 'all' || filterDept !== 'all' || filterEntity !== 'all' || filterChosenApprover !== 'all' || listSearch.trim() !== '') && (
            <button 
              onClick={() => {
                setFilterStatus('all');
                setFilterDept('all');
                setFilterEntity('all');
                setFilterChosenApprover('all');
                setListSearch('');
              }}
              className="flex items-center gap-1 text-[10px] font-bold text-indigo-600 hover:text-indigo-700 bg-indigo-50 px-2 py-1 rounded-full transition-colors"
            >
              <RotateCcw className="w-2.5 h-2.5" />
              Clear Filters
            </button>
          )}
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-zinc-200 overflow-hidden shadow-sm flex-1 min-h-0 flex flex-col min-w-0">
        <div className="overflow-auto flex-1 min-h-0 min-w-0">
          <table className="w-full min-w-[1120px] table-auto border-collapse text-left text-[11px] sm:text-xs">
            <thead className="bg-zinc-50 text-zinc-500 text-[10px] sm:text-xs uppercase font-bold sticky top-0 z-[1] shadow-[0_1px_0_0_rgb(228_228_231)]">
              <tr>
                <th className="px-3 py-2 min-w-[14rem] align-bottom">
                  <button type="button" onClick={() => toggleSort('request')} className="hover:text-zinc-700 transition-colors text-left">
                    Request{sortIndicator('request')}
                  </button>
                </th>
                <th className="px-3 py-2 min-w-[7rem] align-bottom">
                  <button type="button" onClick={() => toggleSort('entity')} className="hover:text-zinc-700 transition-colors text-left">
                    Entity{sortIndicator('entity')}
                  </button>
                </th>
                <th className="px-3 py-2 min-w-[5.5rem] align-bottom">
                  <button type="button" onClick={() => toggleSort('department')} className="hover:text-zinc-700 transition-colors text-left">
                    Dept{sortIndicator('department')}
                  </button>
                </th>
                <th className="px-3 py-2 min-w-[8.5rem] align-bottom">
                  <button type="button" onClick={() => toggleSort('requester')} className="hover:text-zinc-700 transition-colors text-left">
                    Requester{sortIndicator('requester')}
                  </button>
                </th>
                <th className="px-3 py-2 min-w-[8.5rem] align-bottom">
                  <button type="button" onClick={() => toggleSort('chosenApprover')} className="hover:text-zinc-700 transition-colors text-left">
                    Chosen approver{sortIndicator('chosenApprover')}
                  </button>
                </th>
                <th className="px-3 py-2 min-w-[6.5rem] align-bottom">
                  <button type="button" onClick={() => toggleSort('status')} className="hover:text-zinc-700 transition-colors text-left">
                    Status{sortIndicator('status')}
                  </button>
                </th>
                <th className="px-3 py-2 min-w-[6rem] align-bottom text-zinc-500">PO #</th>
                <th className="px-3 py-2 min-w-[7.5rem] align-bottom text-zinc-500">PO status</th>
                <th className="px-3 py-2 min-w-[5.5rem] align-bottom">
                  <button type="button" onClick={() => toggleSort('date')} className="hover:text-zinc-700 transition-colors text-left">
                    Date{sortIndicator('date')}
                  </button>
                </th>
                <th className="px-3 py-2 min-w-[9rem] align-bottom text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-100">
              {sortedFilteredRequests.length === 0 ? (
                <tr>
                  <td colSpan={10} className="px-3 py-12 text-center text-zinc-400 italic">
                    No requests found matching filters.
                  </td>
                </tr>
              ) : (
                sortedFilteredRequests.map((r) => (
                  <tr
                    key={r.id}
                    onClick={() => {
                      void handleViewRequest(r);
                    }}
                    className="hover:bg-zinc-50 transition-colors cursor-pointer align-top"
                  >
                    <td className="px-3 py-2 align-top min-w-[14rem] max-w-md">
                      <p className="font-semibold text-zinc-900 leading-snug break-words">
                        {r.formatted_id ? `[${toUpperSerial(r.formatted_id)}] ${r.title}` : r.title}
                      </p>
                      <p className="text-zinc-500 mt-0.5 break-words text-[10px] sm:text-xs leading-snug">{r.template_name}</p>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[7rem]">
                      <div className="flex flex-col gap-0.5 break-words hyphens-auto leading-tight">
                        <span className="font-bold text-indigo-600 uppercase tracking-wide text-[10px] sm:text-xs">
                          {String(r.entity || '-')}
                        </span>
                        <span className="text-[10px] text-zinc-500">
                          {entityLegalDisplayName(r.entity)}
                        </span>
                      </div>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[5.5rem] text-zinc-700 font-medium break-words leading-tight">
                      {r.department}
                    </td>
                    <td className="px-3 py-2 align-top min-w-[8.5rem] text-zinc-700 break-words leading-tight">{r.requester_name}</td>
                    <td className="px-3 py-2 align-top min-w-[8.5rem] text-zinc-700 break-words leading-tight">
                      {requestHasChosenApprover(r) ? chosenApproverNameLabel(r) : '—'}
                    </td>
                    <td className="px-3 py-2 align-top min-w-[6.5rem]">
                      <span
                        className={cn(
                          'inline-flex text-[9px] sm:text-[10px] uppercase font-bold px-1.5 py-0.5 rounded-full leading-tight',
                          workflowRequestStatusBadgeClass(r)
                        )}
                      >
                        {formatWorkflowRequestStatusLabel(r)}
                      </span>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[6rem] text-zinc-700 font-mono tabular-nums break-all leading-tight">
                      {displayPoNumberInRequestTable(r)}
                    </td>
                    <td className="px-3 py-2 align-top min-w-[7.5rem]">
                      {showPoStatusBadgeInRequestTable(r) ? (
                        <span
                          className={cn(
                            'inline-flex text-[9px] sm:text-[10px] uppercase font-bold px-1.5 py-0.5 rounded-full leading-tight',
                            workflowRequestStatusBadgeClassFromRawStatus(
                              isPO_Only(r) ? r.status : r.linked_po_status
                            )
                          )}
                        >
                          {displayPoStatusLabelInRequestTable(r)}
                        </span>
                      ) : (
                        <span className="text-zinc-400">{displayPoStatusLabelInRequestTable(r)}</span>
                      )}
                    </td>
                    <td className="px-3 py-2 align-top min-w-[5.5rem] text-zinc-500 whitespace-nowrap">
                      {formatDateMYT(r.created_at)}
                    </td>
                    <td className="px-3 py-2 align-top text-right min-w-[9rem]">
                      <div className="flex flex-col items-end gap-1">
                        {canShowConvertPRToPO(r, user) && (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              openConvertPoModal(r);
                            }}
                            title="Convert to PO"
                            className="inline-flex items-center justify-center gap-1 px-2 py-1 bg-indigo-600 text-white rounded-md font-bold text-[10px] hover:bg-indigo-700 transition-colors shadow-sm whitespace-nowrap"
                          >
                            <RefreshCw className="w-3 h-3 shrink-0" />
                            <span className="hidden lg:inline">Convert to PO</span>
                            <span className="lg:hidden">→ PO</span>
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            void handleViewRequest(r);
                          }}
                          title="Request details"
                          className="inline-flex items-center justify-center gap-1 px-2 py-1 bg-indigo-50 text-indigo-600 rounded-md font-bold text-[10px] hover:bg-indigo-100 transition-all border border-indigo-100 shadow-sm whitespace-nowrap"
                        >
                          <FileText className="w-3 h-3 shrink-0" />
                          Details
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {selectedRequest && (
          <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white w-full max-w-[96vw] rounded-2xl shadow-2xl overflow-hidden flex flex-col h-[94vh]"
            >
              <div className="p-6 border-b border-zinc-100 flex items-center justify-between bg-zinc-50 shrink-0">
                <div>
                  <h2 className="text-xl font-bold text-zinc-900">
                    {selectedRequest.formatted_id ? `[${toUpperSerial(selectedRequest.formatted_id)}] ${selectedRequest.title}` : selectedRequest.title}
                  </h2>
                  <p className="text-sm text-zinc-500">Template: {selectedRequest.template_name} • Dept: {selectedRequest.department} • Entity: <span className="font-bold text-indigo-600">{entityLegalDisplayName(selectedRequest.entity)}</span> <span className="text-zinc-400">({selectedRequest.entity?.trim() || '-'})</span></p>
                </div>
                <div className="flex items-center gap-2">
                  {canEdit && !isEditing && (
                    <button 
                      onClick={() => {
                        setEditData({ 
                          title: selectedRequest.title, 
                          details: selectedRequest.details, 
                          line_items: normalizeLineItemsForDateInputs(
                            [...(selectedRequest.line_items || [])],
                            getColumns(selectedRequest)
                          ),
                          tax_rate: procurementUnitRateToPercent(selectedRequest.tax_rate, 0.18),
                          discount_rate: procurementUnitRateToPercent(selectedRequest.discount_rate, 0),
                          currency: selectedRequest.currency?.trim() ?? '',
                          cost_center: selectedRequest.cost_center || '',
                          section: sectionSelectionFromStored(selectedRequest.section),
                          suggested_supplier: prSuggestedSupplierDisplay(selectedRequest),
                          requester_username: String(selectedRequest.requester_username ?? '').trim(),
                          requester_name: String(selectedRequest.requester_name ?? '').trim(),
                        } as any);
                        setEditAttachmentKeep(attachments);
                        setEditAttachmentAdd([]);
                        setIsEditing(true);
                      }}
                      className="flex items-center gap-2 px-3 py-1.5 rounded-lg border border-indigo-200 text-indigo-600 text-xs font-bold hover:bg-indigo-50 transition-all"
                    >
                      <Edit2 className="w-3.5 h-3.5" />
                      Edit Request
                    </button>
                  )}
                  {canShowPurchasingApprovedPRHeaderActions(selectedRequest, user) && !isEditing && (
                    <>
                      <button 
                        onClick={() => handleGeneratePODraft(selectedRequest)}
                        className="flex items-center gap-2 px-4 py-2 bg-amber-50 border border-amber-200 rounded-xl text-sm font-bold text-amber-700 hover:bg-amber-100 transition-colors"
                      >
                        <Download className="w-4 h-4" />
                        Generate PO Draft PDF
                      </button>
                      {canShowConvertPRToPO(selectedRequest, user) && (
                        <button 
                          onClick={() => openConvertPoModal(selectedRequest)}
                          className="flex items-center gap-2 px-4 py-2 bg-indigo-600 rounded-xl text-sm font-bold text-white hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-100"
                        >
                          <RefreshCw className="w-4 h-4" />
                          Convert PR → PO
                        </button>
                      )}
                    </>
                  )}
                  {canShowPurchasingFinalDecision(selectedRequest, user) && !isEditing && (
                    <>
                      <button
                        onClick={() => handlePurchasingFinalDecision(selectedRequest.id, 'rejected')}
                        className="flex items-center gap-2 px-4 py-2 bg-red-50 border border-red-200 rounded-xl text-sm font-bold text-red-700 hover:bg-red-100 transition-colors"
                      >
                        <XCircle className="w-4 h-4" />
                        Reject PR (Resubmittable)
                      </button>
                      <button
                        onClick={() => handlePurchasingFinalDecision(selectedRequest.id, 'cancelled')}
                        className="flex items-center gap-2 px-4 py-2 bg-zinc-200 border border-zinc-300 rounded-xl text-sm font-bold text-zinc-800 hover:bg-zinc-300 transition-colors"
                        title="Cancellation is final and cannot be resubmitted"
                      >
                        <X className="w-4 h-4" />
                        Cancel PR (Final)
                      </button>
                    </>
                  )}
                  {canRequesterCancelSelected && !isEditing && (
                    <button
                      onClick={() => handleRequesterCancel(selectedRequest.id)}
                      disabled={loading}
                      className="flex items-center gap-2 px-4 py-2 bg-zinc-200 border border-zinc-300 rounded-xl text-sm font-bold text-zinc-800 hover:bg-zinc-300 transition-colors disabled:opacity-50"
                      title="Cancel before any approver has approved (final action)"
                    >
                      <X className="w-4 h-4" />
                      Cancel Request
                    </button>
                  )}
                  <button onClick={() => {
                    setSelectedRequest(null);
                    setIsEditing(false);
                    setDetailsViewMode('details');
                    setDetailsPdfPreview((prev) => {
                      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
                      return null;
                    });
                  }} className="text-zinc-400 hover:text-zinc-600">
                    <XCircle className="w-6 h-6" />
                  </button>
                </div>
              </div>
              <div className="p-0 overflow-y-auto flex-1 bg-zinc-50/50">
                {isEditing ? (
                  <div className="p-8 w-full">
                    <div className="bg-white rounded-2xl shadow-sm border border-zinc-200 p-8 space-y-8">
                      <div className="flex items-center gap-3 mb-6">
                        <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center">
                          <Edit2 className="w-5 h-5 text-indigo-600" />
                        </div>
                        <div>
                          <h3 className="text-xl font-bold text-zinc-900">Edit Request</h3>
                          <p className="text-sm text-zinc-500">Update your request details and line items</p>
                        </div>
                      </div>

                      <div className="space-y-6">
                        {isProcurementPRorPORequest(selectedRequest) && (
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Currency</label>
                              <select
                                className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                value={editData.currency}
                                onChange={(e) => setEditData({ ...editData, currency: e.target.value })}
                                required
                              >
                                <option value="">Select currency…</option>
                                {CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">{procurementTaxRateFormLabel(selectedRequest.entity)}</label>
                              <div className="flex items-center gap-3">
                                <input
                                  type="number"
                                  step="0.01"
                                  min="0"
                                  max="100"
                                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                  value={editData.tax_rate !== undefined ? editData.tax_rate : 18}
                                  onChange={(e) => setEditData({ ...editData, tax_rate: parseFloat(e.target.value) || 0 })}
                                />
                                <span className="text-xs text-zinc-500 font-bold uppercase tracking-tighter">%</span>
                              </div>
                            </div>
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Discount (%)</label>
                              <div className="flex items-center gap-3">
                                <input
                                  type="number"
                                  step="0.01"
                                  min="0"
                                  max="100"
                                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                  value={editData.discount_rate !== undefined ? editData.discount_rate : 0}
                                  onChange={(e) => setEditData({ ...editData, discount_rate: parseFloat(e.target.value) || 0 })}
                                />
                                <span className="text-xs text-zinc-500 font-bold uppercase tracking-tighter">%</span>
                              </div>
                            </div>
                          </div>
                        )}
                        {isSRRequest(selectedRequest) && (
                          <div className="max-w-md">
                            <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Section</label>
                            <select
                              className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                              value={sectionSelectionFromStored(editData.section)}
                              onChange={(e) => setEditData({ ...editData, section: e.target.value })}
                              required
                            >
                              <option value="" disabled>Select section...</option>
                              {SECTION_OPTIONS.map((opt) => (
                                <option key={opt} value={opt}>{opt}</option>
                              ))}
                            </select>
                          </div>
                        )}
                        {isPRRequest(selectedRequest) && (
                          <div className="max-w-2xl">
                            <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Suggested supplier</label>
                            <input
                              type="text"
                              className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                              value={editData.suggested_supplier ?? ''}
                              onChange={(e) => setEditData({ ...editData, suggested_supplier: e.target.value })}
                              required
                            />
                            <p className="text-xs text-zinc-500 mt-1">One supplier for the entire PR.</p>
                          </div>
                        )}
                        <div>
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Request Title</label>
                          <input
                            type="text"
                            className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                            value={editData.title}
                            onChange={(e) => setEditData({ ...editData, title: e.target.value })}
                          />
                        </div>
                        {isPO_Only(selectedRequest) && isPurchasing && isWorkflowRequestPending(selectedRequest) && (
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 max-w-3xl">
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Requester login</label>
                              <input
                                type="text"
                                className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                autoComplete="off"
                                placeholder="Username in system"
                                value={editData.requester_username}
                                onChange={(e) => setEditData({ ...editData, requester_username: e.target.value })}
                              />
                            </div>
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Requester name</label>
                              <input
                                type="text"
                                className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                placeholder="Name on document"
                                value={editData.requester_name}
                                onChange={(e) => setEditData({ ...editData, requester_name: e.target.value })}
                              />
                            </div>
                            <p className="text-xs text-zinc-500 sm:col-span-2">
                              Changing login reassigns the requester and sets department from their profile. Changing only the name updates department from the profile of the current requester.
                            </p>
                          </div>
                        )}
                        {!isProcurementPRorPORequest(selectedRequest) && (
                        <div>
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Details</label>
                          <textarea
                            rows={4}
                            className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                            value={editData.details}
                            onChange={(e) => setEditData({ ...editData, details: e.target.value })}
                          />
                        </div>
                        )}

                        <div>
                          <div className="flex items-center justify-between mb-4">
                            <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">
                              {isPRRequest(selectedRequest) ? 'PR line items' : isSRRequest(selectedRequest) ? 'Stock requisition line items' : isPO_Only(selectedRequest) ? 'PO line items' : 'Line items'}
                            </label>
                            <button
                              onClick={() => {
                                const newItem: any = { [LINE_ITEM_REMARKS_KEY]: '' };
                                const columns = getColumns(selectedRequest);
                                columns.forEach(col => newItem[col] = '');
                                setEditData({ ...editData, line_items: [...editData.line_items, newItem] });
                              }}
                              className="flex items-center gap-1 text-xs font-bold text-indigo-600 hover:text-indigo-800"
                            >
                              <Plus className="w-3 h-3" />
                              Add Item
                            </button>
                          </div>
                          <div className="overflow-x-auto border border-zinc-200 rounded-xl">
                            <table className="w-full text-sm">
                              <thead className="bg-zinc-50 border-b border-zinc-200">
                                <tr>
                                  <th className="px-3 py-2 font-bold text-zinc-600 w-10 text-center">#</th>
                                  {getColumns(selectedRequest).map(col => (
                                    <th key={col} className="px-3 py-2 font-bold text-zinc-600 text-left">{col}</th>
                                  ))}
                                  <th className="px-3 py-2 font-bold text-zinc-600 w-10 text-center"></th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-zinc-200">
                                {editData.line_items.map((item, idx) => {
                                  const cols = getColumns(selectedRequest);
                                  return (
                                  <React.Fragment key={idx}>
                                  <tr className="group hover:bg-zinc-50 transition-colors">
                                    <td className="px-3 py-2 text-center text-zinc-400 font-bold">{idx + 1}</td>
                                    {cols.map(col => (
                                      <td key={col} className="px-3 py-2">
                                        {isCostCenterGridColumn(col) ? (
                                          <SearchableSelect
                                            options={costCenterOptions}
                                            value={item[col] || ''}
                                            onChange={(val) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = { ...newItems[idx], [col]: val };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                            placeholder="Select Cost Center..."
                                            loading={costCentersLoading}
                                            loadFailed={costCentersLoadFailed}
                                            loadErrorHint={costCentersFetchError}
                                          />
                                        ) : (isSRRequest(selectedRequest) && isSpareLocationColumn(col)) ? (
                                          <SearchableSelect
                                            options={costCenterOptions}
                                            value={item[col] || ''}
                                            onChange={(val) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = { ...newItems[idx], [col]: val };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                            placeholder="Select Spare for (Location)..."
                                            loading={costCentersLoading}
                                            loadFailed={costCentersLoadFailed}
                                            loadErrorHint={costCentersFetchError}
                                          />
                                        ) : isProcurementLineItemDateColumn(col) ? (
                                          <input
                                            type="date"
                                            className="w-full px-2 py-1 rounded border border-zinc-200 text-xs outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
                                            value={htmlDateValueFromStored(item[col])}
                                            onChange={(e) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = { ...newItems[idx], [col]: e.target.value };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                          />
                                        ) : (
                                          <input
                                            type="text"
                                            className="w-full px-2 py-1 rounded border border-zinc-200 text-xs outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
                                            value={item[col] || ''}
                                            onChange={(e) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = { ...newItems[idx], [col]: e.target.value };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                          />
                                        )}
                                      </td>
                                    ))}
                                    <td className="px-3 py-2 text-center">
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const newItems = [...editData.line_items];
                                          newItems.splice(idx, 1);
                                          setEditData({ ...editData, line_items: newItems });
                                        }}
                                        className="text-zinc-300 hover:text-red-500 transition-colors"
                                      >
                                        <X className="w-4 h-4" />
                                      </button>
                                    </td>
                                  </tr>
                                  <tr className="bg-zinc-50/90">
                                    <td colSpan={cols.length + 2} className="px-3 py-2">
                                      <label className="text-[10px] font-bold text-zinc-500 uppercase">Line remarks (optional)</label>
                                      <textarea
                                        rows={2}
                                        className="w-full mt-1 px-2 py-1.5 rounded border border-zinc-200 text-xs outline-none focus:ring-1 focus:ring-indigo-500 resize-y min-h-[52px]"
                                        value={item[LINE_ITEM_REMARKS_KEY] ?? ''}
                                        onChange={(e) => {
                                          const newItems = [...editData.line_items];
                                          newItems[idx] = { ...newItems[idx], [LINE_ITEM_REMARKS_KEY]: e.target.value };
                                          setEditData({ ...editData, line_items: newItems });
                                        }}
                                        placeholder="Extra notes for this line…"
                                      />
                                    </td>
                                  </tr>
                                  </React.Fragment>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        <div className="space-y-3">
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block">Attachments</label>
                          <div className="flex flex-wrap gap-2">
                            {editAttachmentKeep.map((att) => (
                              <div key={`keep-${att.id}`} className="flex items-center gap-2 bg-zinc-100 px-3 py-1 rounded-full text-xs text-zinc-600 border border-zinc-200">
                                <Paperclip className="w-3 h-3" />
                                <span>{att.file_name}</span>
                                <button
                                  type="button"
                                  onClick={() => setEditAttachmentKeep((prev) => prev.filter((a) => a.id !== att.id))}
                                  title="Remove existing attachment"
                                >
                                  <Trash2 className="w-3 h-3 hover:text-red-500" />
                                </button>
                              </div>
                            ))}
                            {editAttachmentAdd.map((att) => (
                              <div key={att.id} className="flex items-center gap-2 bg-indigo-50 px-3 py-1 rounded-full text-xs text-indigo-700 border border-indigo-100">
                                <Upload className="w-3 h-3" />
                                <span>{att.name}</span>
                                <button
                                  type="button"
                                  onClick={() => setEditAttachmentAdd((prev) => prev.filter((a) => a.id !== att.id))}
                                  title="Remove new attachment"
                                >
                                  <Trash2 className="w-3 h-3 hover:text-red-500" />
                                </button>
                              </div>
                            ))}
                          </div>
                          <label className="flex flex-col items-center justify-center w-full h-20 border-2 border-dashed border-zinc-300 rounded-xl cursor-pointer hover:bg-zinc-50 transition-colors">
                            <div className="flex flex-col items-center justify-center">
                              <Upload className="w-5 h-5 text-zinc-400 mb-1" />
                              <p className="text-xs text-zinc-500">Re-upload or add more files</p>
                            </div>
                            <input type="file" className="hidden" multiple onChange={handleEditAttachmentAdd} />
                          </label>
                        </div>

                        <div className="flex gap-3 pt-6 border-t border-zinc-100">
                          <button
                            onClick={handleSaveEdit}
                            disabled={loading}
                            className="flex-1 bg-indigo-600 text-white py-3 rounded-xl font-bold text-sm hover:bg-indigo-700 transition-all disabled:opacity-50"
                          >
                            Save Changes
                          </button>
                          <button
                            onClick={() => setIsEditing(false)}
                            className="flex-1 bg-zinc-100 text-zinc-600 py-3 rounded-xl font-bold text-sm hover:bg-zinc-200 transition-all"
                          >
                            Cancel
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="w-full bg-white my-8 shadow-sm border border-zinc-200 rounded-xl overflow-hidden">
                    {/* Document Header */}
                    <div className="p-8 border-b border-zinc-100 bg-zinc-50/30">
                      <div className="flex justify-between items-start mb-8">
                        <div>
                          <h1 className="text-2xl font-black text-zinc-900 uppercase tracking-tight">
                            {isPRRequest(selectedRequest)
                              ? 'Purchase Requisition Form'
                              : isSRRequest(selectedRequest)
                                ? 'Stock Item Requisition Form'
                                : isPO_Only(selectedRequest)
                                  ? 'Purchase Order'
                                  : 'Workflow Request'}
                          </h1>
                          <p className="text-sm text-zinc-500 font-medium">Ref: #{displayRequestSerial(selectedRequest)}</p>
                        </div>
                        <div className="flex items-center gap-4">
                          <button 
                            onClick={() => handlePrintPR(selectedRequest)}
                            className="inline-flex items-center gap-2 px-3 py-2 rounded-lg border border-zinc-200 text-zinc-600 bg-white hover:bg-zinc-50 hover:text-indigo-600 transition-colors"
                            title={isPR_Only(selectedRequest) ? 'View PR form' : isSR_Only(selectedRequest) ? 'View SR form' : 'View PDF'}
                          >
                            <FileText className="w-4 h-4" />
                            <span className="text-xs font-bold">
                              {isPR_Only(selectedRequest) ? 'View PR form' : isSR_Only(selectedRequest) ? 'View SR form' : 'View PDF'}
                            </span>
                          </button>
                          <div className="text-right">
                            <span className={cn(
                              "px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider",
                              workflowRequestStatusBadgeClass(selectedRequest)
                            )}>
                              {formatWorkflowRequestStatusLabel(selectedRequest)}
                            </span>
                            <p className="text-[10px] text-zinc-400 mt-2 font-bold uppercase tracking-widest">
                              {formatDateMYT(selectedRequest.created_at)}
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Summary Cards */}
                      <div className="grid grid-cols-4 gap-4 mb-8">
                        <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Total Amount</p>
                          <p className="text-xl font-black text-indigo-600">
                            {procurementCurrencyPrefix(selectedRequest.currency)}
                            {formatProcurementMoney(
                              isProcurementPRorPORequest(selectedRequest)
                                ? procurementMoneyTotals(selectedRequest).total
                                : ((selectedRequest.line_items?.reduce((sum, item) => sum + (parseFloat(item['Quantity'] || '0') * parseFloat(item['Unit Price'] || '0')), 0) || 0) *
                                    (1 + (selectedRequest.tax_rate !== undefined ? selectedRequest.tax_rate : 0.18)))
                            )}
                          </p>
                        </div>
                        <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Items Count</p>
                          <p className="text-xl font-black text-zinc-900">{selectedRequest.line_items?.length || 0} Items</p>
                        </div>
                        <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Requester</p>
                          <p className="text-sm font-bold text-zinc-900 truncate">{selectedRequest.requester_name}</p>
                          <p className="text-[10px] text-zinc-400 font-bold uppercase">{selectedRequest.department}</p>
                        </div>
                        <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Current Step</p>
                          <p className="text-sm font-bold text-zinc-900">
                            {isWorkflowRequestPending(selectedRequest) 
                              ? selectedRequest.template_steps[selectedRequest.current_step_index]?.label || 'Processing'
                              : formatWorkflowRequestStatusLabel(selectedRequest)}
                          </p>
                          <p className="text-[10px] text-zinc-400 font-bold uppercase">Workflow Progress</p>
                        </div>
                      </div>

                    <div className="grid grid-cols-2 gap-8">
                      <div>
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Requester Information</label>
                        <p className="font-bold text-zinc-900">{selectedRequest.requester_name}</p>
                        <p className="text-sm text-zinc-500">{selectedRequest.department}</p>
                      </div>
                      <div className="text-right">
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Template Type</label>
                        <p className="font-bold text-zinc-900">{selectedRequest.template_name}</p>
                        <p className="text-sm text-zinc-500">Category: {selectedRequest.category || 'General'}</p>
                      </div>
                    </div>

                    {requestHasChosenApprover(selectedRequest) ? (
                      <div className="mt-6 pt-6 border-t border-zinc-100">
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Chosen approver</label>
                        <p className="font-bold text-zinc-900">{chosenApproverNameLabel(selectedRequest)}</p>
                        {String(selectedRequest.assigned_approver_designation ?? '').trim() ? (
                          <p className="text-sm text-zinc-500">{String(selectedRequest.assigned_approver_designation).trim()}</p>
                        ) : null}
                      </div>
                    ) : null}

                    {isProcurementPRorPORequest(selectedRequest) && (
                      <div className="grid grid-cols-2 gap-8 mt-8 pt-8 border-t border-zinc-100">
                        <div>
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Currency</label>
                          <p className="font-bold text-zinc-900">{procurementCurrencyLabel(selectedRequest.currency)}</p>
                        </div>
                        <div className="text-right">
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Cost Center</label>
                          <p className="font-bold text-zinc-900">
                            {String(selectedRequest.cost_center || '').trim()
                              ? formatCostCenterCatalogCell(selectedRequest.cost_center, costCenterOptions)
                              : 'Not Specified'}
                          </p>
                          {isSRRequest(selectedRequest) && (
                            <>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mt-3 mb-1">Section</label>
                              <p className="font-bold text-zinc-900">{selectedRequest.section || 'Not Specified'}</p>
                            </>
                          )}
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="p-8 space-y-10">
                    <section>
                      <div className="flex items-center justify-between gap-3 mb-3 border-b border-zinc-100 pb-2">
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">Details</h4>
                        <div className="inline-flex rounded-lg border border-zinc-200 overflow-hidden">
                          <button
                            type="button"
                            onClick={() => setDetailsViewMode('details')}
                            className={cn(
                              "px-3 py-1.5 text-[10px] font-black uppercase tracking-widest transition-colors",
                              detailsViewMode === 'details'
                                ? "bg-indigo-600 text-white"
                                : "bg-white text-zinc-500 hover:bg-zinc-50"
                            )}
                          >
                            Details
                          </button>
                          <button
                            type="button"
                            onClick={() => handleShowRequestPdfInline(selectedRequest)}
                            className={cn(
                              "px-3 py-1.5 text-[10px] font-black uppercase tracking-widest transition-colors border-l border-zinc-200 inline-flex items-center gap-1.5",
                              detailsViewMode === 'pdf'
                                ? "bg-indigo-600 text-white"
                                : "bg-white text-zinc-500 hover:bg-zinc-50"
                            )}
                          >
                            <FileText className="w-3 h-3" />
                            PDF
                          </button>
                        </div>
                      </div>
                      {detailsViewMode === 'pdf' ? (
                        <div className="w-full h-[72vh] border border-zinc-200 rounded-xl overflow-hidden bg-zinc-100">
                          {detailsPdfPreview ? (
                            <iframe
                              src={`${detailsPdfPreview.url}#view=FitH`}
                              className="w-full h-full border-none"
                              title="Request PDF Preview"
                            />
                          ) : (
                            <div className="h-full flex items-center justify-center text-sm text-zinc-500">
                              PDF is not ready.
                            </div>
                          )}
                        </div>
                      ) : (
                        !isProcurementPRorPORequest(selectedRequest) ? (
                          <p className="text-zinc-700 whitespace-pre-wrap leading-relaxed">{selectedRequest.details}</p>
                        ) : (
                          <p className="text-zinc-500 text-sm">Switch to PDF to view the full form layout.</p>
                        )
                      )}
                    </section>

                    {isPRRequest(selectedRequest) && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 border-b border-zinc-100 pb-2">
                          Suggested supplier
                        </h4>
                        <p className="text-sm text-zinc-800 font-medium">
                          {prSuggestedSupplierDisplay(selectedRequest) || '—'}
                        </p>
                      </section>
                    )}

                    {/* Line Items */}
                    {selectedRequest.line_items && selectedRequest.line_items.length > 0 && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-3 border-b border-zinc-100 pb-2">
                          {isPRRequest(selectedRequest) ? 'PR line items' : isSRRequest(selectedRequest) ? 'Stock requisition line items' : isPO_Only(selectedRequest) ? 'PO line items' : 'Line items'}
                        </h4>
                        <div className="border border-zinc-200 rounded-lg overflow-hidden">
                          <table className="w-full text-sm">
                            <thead className="bg-zinc-50 border-b border-zinc-200">
                              <tr>
                                <th className="px-4 py-2 font-bold text-zinc-600 w-12 text-center">No</th>
                                {procurementGridColumns(selectedRequest).map(col => (
                                  <th key={col} className="px-4 py-2 font-bold text-zinc-600 text-left">{col}</th>
                                ))}
                                {procurementRowShowsLineTotal(selectedRequest) && <th className="px-4 py-2 font-bold text-zinc-600 text-right">Total</th>}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-zinc-200">
                              {selectedRequest.line_items.map((item, idx) => {
                                const cols = procurementGridColumns(selectedRequest);
                                const showLineTotal = procurementRowShowsLineTotal(selectedRequest);
                                const qty = lineItemQtyForDisplay(item);
                                const price = parseFloat(item['Unit Price'] || '0');
                                const total = qty * price;
                                const lineRem = item[LINE_ITEM_REMARKS_KEY] && String(item[LINE_ITEM_REMARKS_KEY]).trim();
                                return (
                                  <React.Fragment key={item.id || idx}>
                                  <tr className={idx % 2 === 0 ? 'bg-white' : 'bg-zinc-50/30'}>
                                    <td className="px-4 py-3 text-center text-zinc-500 font-medium">{idx + 1}</td>
                                    {cols.map(col => (
                                      <td key={col} className="px-4 py-3 text-zinc-700">
                                        {(col === 'Unit Price' || col === 'Price' || col === 'Amount')
                                          ? `${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(Number(item[col] || 0))}`.trim()
                                          : (col === REMARKS_LINE_COL
                                            ? (lineItemRemarksDisplay(item) || '-')
                                            : (isCostCenterGridColumn(col) || (isSRRequest(selectedRequest) && isSpareLocationColumn(col)))
                                              ? formatCostCenterCatalogCell(item[col], costCenterOptions)
                                              : (item[col] || '-'))}
                                      </td>
                                    ))}
                                    {showLineTotal && (
                                      <td className="px-4 py-3 text-right font-bold text-zinc-900">
                                        {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(total)}`.trim()}
                                      </td>
                                    )}
                                  </tr>
                                  {lineRem && (
                                    <tr className="bg-indigo-50/40 border-t border-indigo-100/80">
                                      <td colSpan={cols.length + 1 + (showLineTotal ? 1 : 0)} className="px-4 py-2.5 text-xs text-zinc-700">
                                        <span className="font-bold text-zinc-500 uppercase tracking-wide text-[10px]">Line remarks: </span>
                                        <span className="whitespace-pre-wrap">{lineRem}</span>
                                      </td>
                                    </tr>
                                  )}
                                  </React.Fragment>
                                );
                              })}
                            </tbody>
                            {procurementRowShowsLineTotal(selectedRequest) && (() => {
                              const m = procurementMoneyTotals(selectedRequest);
                              const gc = procurementGridColumns(selectedRequest);
                              return (
                              <tfoot className="bg-zinc-50/50 font-bold">
                                <tr>
                                  <td colSpan={gc.length + 1} className="px-4 py-2 text-right text-zinc-500">Subtotal</td>
                                  <td className="px-4 py-2 text-right text-zinc-900">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.subtotal)}`.trim()}
                                  </td>
                                </tr>
                                <tr>
                                  <td colSpan={gc.length + 1} className="px-4 py-2 text-right text-zinc-500">Discount ({(m.discountRate * 100).toFixed(0)}%)</td>
                                  <td className="px-4 py-2 text-right text-zinc-900">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.discountAmount)}`.trim()}
                                  </td>
                                </tr>
                                <tr>
                                  <td colSpan={gc.length + 1} className="px-4 py-2 text-right text-zinc-500">{procurementTaxLabelForEntity(selectedRequest.entity)} ({(m.taxRate * 100).toFixed(0)}%)</td>
                                  <td className="px-4 py-2 text-right text-zinc-900">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.taxAmount)}`.trim()}
                                  </td>
                                </tr>
                                <tr className="bg-zinc-100/50 text-lg">
                                  <td colSpan={gc.length + 1} className="px-4 py-4 text-right text-zinc-900 uppercase tracking-tight">Total Amount</td>
                                  <td className="px-4 py-4 text-right text-indigo-600 font-black">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.total)}`.trim()}
                                  </td>
                                </tr>
                              </tfoot>
                              );
                            })()}
                          </table>
                        </div>
                      </section>
                    )}

                    {isPO_Only(selectedRequest) && isPurchasing && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-3 border-b border-zinc-100 pb-2">
                          Real Purchase Order
                        </h4>
                        <label className="inline-flex items-center gap-2 px-3 py-2 rounded-lg border border-zinc-200 bg-zinc-50 hover:bg-zinc-100 cursor-pointer text-xs font-semibold text-zinc-700 transition-colors">
                          <Upload className="w-4 h-4" />
                          {loading ? 'Uploading…' : 'Upload PO file (optional)'}
                          <input
                            type="file"
                            className="hidden"
                            disabled={loading}
                            onChange={(e) => {
                              const f = e.target.files?.[0];
                              if (!f) return;
                              void handleUploadRealPoAttachment(f);
                              e.currentTarget.value = '';
                            }}
                          />
                        </label>
                        <p className="text-[11px] text-zinc-500 mt-2">
                          You can upload the official PO now or later.
                        </p>
                      </section>
                    )}

                    {/* Attachments */}
                    {attachments.length > 0 && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-3 border-b border-zinc-100 pb-2">Attachments</h4>
                        <div className="grid grid-cols-2 gap-3">
                          {attachments.map((att) => (
                            <div
                              key={att.id}
                              onClick={() => {
                                void openWorkflowRequestAttachment(selectedRequest.id, att, setViewingPdf);
                              }}
                              className="flex items-center gap-3 p-3 bg-zinc-50 rounded-xl border border-zinc-200 hover:border-indigo-300 transition-all group cursor-pointer"
                            >
                              <div className="w-8 h-8 bg-white rounded-lg flex items-center justify-center border border-zinc-100 group-hover:bg-indigo-50 transition-colors">
                                <FileText className="w-4 h-4 text-indigo-500" />
                              </div>
                              <div className="flex-1 min-w-0">
                                <p className="text-xs font-bold text-zinc-700 truncate">{att.file_name}</p>
                                <p className="text-[10px] text-zinc-400 uppercase font-bold">
                                  {(att.file_type === 'application/pdf' || att.file_name.toLowerCase().endsWith('.pdf')) ? 'View PDF' : 'Download File'}
                                </p>
                              </div>
                              {(att.file_type === 'application/pdf' || att.file_name.toLowerCase().endsWith('.pdf')) ? (
                                <Eye className="w-3 h-3 text-zinc-300 group-hover:text-indigo-500" />
                              ) : (
                                <Download className="w-3 h-3 text-zinc-300 group-hover:text-indigo-500" />
                              )}
                            </div>
                          ))}
                        </div>
                      </section>
                    )}

                    {/* Signatures & Approvals */}
                    <section className="pt-8 border-t-2 border-dashed border-zinc-100">
                      <div className="grid grid-cols-2 gap-12">
                        {/* Requester Signature */}
                        <div>
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-4">Requested By</p>
                          {hasRequesterSignatureProof(selectedRequest) ? (
                            <div className="space-y-2">
                              {isSignatureImageDataUrl(selectedRequest.requester_signature) ? (
                                <img src={selectedRequest.requester_signature!} alt="Requester Signature" className="h-16 object-contain" referrerPolicy="no-referrer" />
                              ) : (
                                <div className="rounded-lg border border-emerald-200 bg-emerald-50/80 px-3 py-2">
                                  <p className="text-xs font-semibold text-emerald-800">Electronic signature</p>
                                  {selectedRequest.requester_signed_at && (
                                    <p className="text-[10px] text-emerald-700/90 mt-0.5">{formatSignatureProofTimestamp(selectedRequest.requester_signed_at)}</p>
                                  )}
                                </div>
                              )}
                              <div className="pt-2 border-t border-zinc-200">
                                <p className="text-sm font-bold text-zinc-900">{selectedRequest.requester_name}</p>
                                <p className="text-xs text-zinc-500">{selectedRequest.department}</p>
                                <p className="text-[10px] text-zinc-400 mt-1">{formatDateMYT(selectedRequest.created_at)}</p>
                              </div>
                            </div>
                          ) : (
                            <div className="h-24 border-2 border-dashed border-zinc-100 rounded-xl flex items-center justify-center">
                              <p className="text-xs text-zinc-300 italic">No signature</p>
                            </div>
                          )}
                        </div>

                        {/* Approvals */}
                        <div className="space-y-8">
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-4">Approval Workflow</p>
                          {selectedRequest.template_steps.map((step, i) => {
                            const approval = [...approvals].reverse().find((a) => a.step_index === i);
                            const isCurrent = isWorkflowRequestPending(selectedRequest) && selectedRequest.current_step_index === i;
                            
                            return (
                              <div key={step.id} className="relative pl-8 border-l-2 border-zinc-100 pb-8 last:pb-0">
                                <div className={cn(
                                  "absolute -left-[9px] top-0 w-4 h-4 rounded-full border-2 border-white",
                                  approval ? (approval.status === 'approved' ? "bg-emerald-500" : "bg-red-500") :
                                  isCurrent ? "bg-amber-500 animate-pulse" : "bg-zinc-200"
                                )} />
                                
                                <div className="space-y-2">
                                  <div className="flex justify-between items-start">
                                    <div>
                                      <p className="text-xs font-black text-zinc-900 uppercase tracking-tight">{step.label}</p>
                                      <p className="text-[10px] text-zinc-400 font-bold uppercase">Role: {step.approverRole}</p>
                                    </div>
                                    {approval && (
                                      <span className={cn(
                                        "text-[10px] font-black uppercase px-2 py-0.5 rounded",
                                        approval.status === 'approved' ? "bg-emerald-100 text-emerald-700" : "bg-red-100 text-red-700"
                                      )}>
                                        {approval.status}
                                      </span>
                                    )}
                                  </div>

                                  {approval && (
                                    <div className="mt-4 p-4 bg-zinc-50 rounded-xl border border-zinc-100">
                                      {approval.comment && <p className="text-xs text-zinc-600 italic mb-3">"{approval.comment}"</p>}
                                      <div className="flex items-center gap-3">
                                        {isSignatureImageDataUrl(approval.approver_signature) ? (
                                          <img src={approval.approver_signature!} alt="Approver Signature" className="h-10 object-contain" referrerPolicy="no-referrer" />
                                        ) : approval.status.toLowerCase() === 'approved' ? (
                                          <div className="shrink-0 rounded-lg border border-emerald-200 bg-emerald-50/80 px-2 py-1.5">
                                            <p className="text-[10px] font-semibold text-emerald-800">E-signed</p>
                                            <p className="text-[9px] text-emerald-700/90">{formatSignatureProofTimestamp(approval.created_at)}</p>
                                          </div>
                                        ) : null}
                                        <div>
                                          <p className="text-xs font-bold text-zinc-900">{approval.approver_name}</p>
                                          {approval.signed_by_name ? (
                                            <p className="text-[10px] text-indigo-600 font-medium">
                                              Signed by proxy: {approval.signed_by_name}
                                            </p>
                                          ) : null}
                                          <p className="text-[10px] text-zinc-400">{formatDateMYT(approval.created_at)}</p>
                                        </div>
                                      </div>
                                    </div>
                                  )}

                                  {isCurrent && (
                                    <div className="mt-4 p-4 bg-amber-50 border border-amber-200 rounded-xl">
                                      <p className="text-xs font-bold text-amber-700 flex items-center gap-2">
                                        <Clock className="w-3 h-3" />
                                        Awaiting decision...
                                      </p>
                                    </div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </section>

                    {/* Approval Form (Integrated) */}
                    {isWorkflowRequestPending(selectedRequest) && canApprove && (
                      <section className="pt-12 border-t-2 border-zinc-100">
                        <div className="bg-indigo-50/50 p-8 rounded-2xl border border-indigo-100">
                          <div className="flex items-center gap-3 mb-6">
                            <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shadow-lg shadow-indigo-200">
                              <Shield className="w-5 h-5 text-white" />
                            </div>
                            <div>
                              <h3 className="text-lg font-black text-zinc-900 uppercase tracking-tight">Your Decision</h3>
                              <p className="text-xs text-zinc-500">Review the document above before proceeding</p>
                            </div>
                          </div>

                          <div className="space-y-6">
                            {mustPickOnBehalf && (
                              <div>
                                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">
                                  Sign on behalf of <span className="text-red-500 ml-1">* Required</span>
                                </label>
                                <select
                                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
                                  value={onBehalfApproverId === '' ? '' : String(onBehalfApproverId)}
                                  onChange={(e) => {
                                    const v = e.target.value;
                                    setOnBehalfApproverId(v === '' ? '' : Number(v));
                                  }}
                                  disabled={hasFixedAssignedApprover || onBehalfLoading || onBehalfOptions.length === 0}
                                >
                                  <option value="">
                                    {onBehalfLoading ? 'Loading approvers…' : 'Select approver'}
                                  </option>
                                  {onBehalfOptions.map((u) => (
                                    <option key={u.id} value={u.id}>
                                      {u.username}
                                      {u.department ? ` — ${u.department}` : ''}
                                    </option>
                                  ))}
                                </select>
                                {hasFixedAssignedApprover ? (
                                  <p className="text-[11px] text-zinc-500 mt-1">
                                    Fixed by requester selection for this PR.
                                  </p>
                                ) : null}
                                {!onBehalfLoading && onBehalfOptions.length === 0 ? (
                                  <p className="text-xs text-amber-700 mt-1">
                                    No eligible approvers for this request.
                                  </p>
                                ) : null}
                              </div>
                            )}
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Comment (Optional)</label>
                              <textarea
                                rows={3}
                                className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
                                placeholder="Add a reason for your decision..."
                                value={comment}
                                onChange={(e) => setComment(e.target.value)}
                              />
                            </div>

                            {requiresProcurementApproverSignaturePad(selectedRequest) && (
                              <div>
                                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">
                                  Signature <span className="text-red-500 ml-1">* Required</span>
                                </label>
                                <div className="bg-white rounded-xl border border-zinc-200 overflow-hidden">
                                  <SignaturePad
                                    onSave={setApproverSignature}
                                    onClear={() => setApproverSignature(null)}
                                    value={approverSignature}
                                    savedSignature={savedApproverSignature}
                                    onUseSaved={useSavedApproverSignature}
                                    onSaveDefault={savedApproverSignatureLoading ? undefined : saveApproverSignatureAsDefault}
                                    onClearSaved={savedApproverSignatureLoading ? undefined : removeSavedApproverSignature}
                                  />
                                </div>
                                {savedApproverSignatureLoading ? (
                                  <p className="text-[11px] text-zinc-400 mt-1">Syncing saved signature…</p>
                                ) : null}
                              </div>
                            )}

                            <div className="flex gap-4 pt-2">
                              <button
                                onClick={() => handleApprove(selectedRequest.id, 'approved')}
                                disabled={loading}
                                className="flex-1 bg-emerald-600 text-white py-4 rounded-xl font-black uppercase tracking-widest hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-100 flex items-center justify-center gap-2 disabled:opacity-50"
                              >
                                <CheckCircle2 className="w-5 h-5" />
                                Approve Document
                              </button>
                              <button
                                onClick={() => handleApprove(selectedRequest.id, 'rejected')}
                                disabled={loading}
                                className="flex-1 bg-red-600 text-white py-4 rounded-xl font-black uppercase tracking-widest hover:bg-red-700 transition-all shadow-lg shadow-red-100 flex items-center justify-center gap-2 disabled:opacity-50"
                              >
                                <XCircle className="w-5 h-5" />
                                Reject Document
                              </button>
                            </div>
                          </div>
                        </div>
                      </section>
                    )}

                    {isWorkflowRequestPending(selectedRequest) && !canApprove && (
                      <div className="mt-8 p-4 bg-zinc-100 rounded-xl border border-zinc-200 flex items-center gap-3 text-zinc-500 text-sm">
                        <Shield className="w-5 h-5" />
                        <span>You don't have the required role ({currentStep?.approverRole}) to approve this step.</span>
                      </div>
                    )}

                    {isWorkflowRequestRejected(selectedRequest) && user.id === selectedRequest.requester_id && (
                      <section className="pt-12 border-t-2 border-zinc-100">
                        <div className="bg-rose-50 p-8 rounded-2xl border border-rose-200">
                          <div className="flex items-center gap-3 mb-6">
                            <div className="w-10 h-10 bg-rose-600 rounded-xl flex items-center justify-center shadow-lg shadow-rose-200">
                              <RefreshCw className="w-5 h-5 text-white" />
                            </div>
                            <div>
                              <h3 className="text-lg font-black text-zinc-900 uppercase tracking-tight">Resubmit Request</h3>
                              <p className="text-xs text-zinc-500">This request was rejected. You can resubmit it to restart the workflow.</p>
                            </div>
                          </div>
                          {(() => {
                            const ra = purchasingRejectApprovalForRequest({ ...selectedRequest, approvals });
                            const reason = String(ra?.comment || '').trim();
                            const who = String(ra?.approver_name || '').trim();
                            const when = ra?.created_at ? formatDateTimeMYT(ra.created_at) : '';
                            if (!reason && !who && !when) return null;
                            return (
                              <div className="mb-6 bg-rose-100 border border-rose-300 rounded-xl p-4">
                                <div className="text-[10px] font-black text-rose-900 uppercase tracking-widest mb-1">
                                  Rejected by purchasing
                                </div>
                                <div className="text-xs text-rose-900">
                                  {who ? <span className="font-bold">{who}</span> : null}
                                  {when ? <span className="text-rose-800/80">{who ? ` • ${when}` : when}</span> : null}
                                </div>
                                {reason ? (
                                  <div className="mt-2 text-sm text-rose-950 whitespace-pre-wrap break-words">“{reason}”</div>
                                ) : null}
                              </div>
                            );
                          })()}
                          <button
                            onClick={() => handleResubmit(selectedRequest.id)}
                            disabled={loading}
                            className="w-full bg-indigo-600 text-white py-4 rounded-xl font-black uppercase tracking-widest hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 flex items-center justify-center gap-2 disabled:opacity-50"
                          >
                            <Send className="w-5 h-5" />
                            Resubmit Now
                          </button>
                        </div>
                      </section>
                    )}
                    {isWorkflowRequestCancelled(selectedRequest) && (
                      <section className="pt-12 border-t-2 border-zinc-100">
                        <div className="bg-rose-50 p-6 rounded-2xl border border-rose-200">
                          {(() => {
                            const requesterCancel = requesterCancelApprovalForRequest({ ...selectedRequest, approvals });
                            const purchasingCancel = purchasingCancelApprovalForRequest({ ...selectedRequest, approvals });
                            const ca = requesterCancel || purchasingCancel;
                            const reason = String(ca?.comment || '').trim();
                            const who = String(ca?.approver_name || '').trim();
                            const when = ca?.created_at ? formatDateTimeMYT(ca.created_at) : '';
                            const byRequester = !!requesterCancel;
                            return (
                              <div className="space-y-2">
                                <p className="text-sm font-bold text-rose-800">
                                  {byRequester ? 'Cancelled by requester' : 'Cancelled by purchasing team'}
                                  {who ? ` — ${who}` : ''}{when ? ` (${when})` : ''}. This is final (cannot be resubmitted).
                                </p>
                                {reason ? (
                                  <div className="text-sm text-rose-950 bg-rose-100 border border-rose-300 rounded-xl p-3">
                                    <div className="text-[10px] font-black text-rose-900 uppercase tracking-widest mb-1">Cancellation reason</div>
                                    <div className="whitespace-pre-wrap break-words">{reason}</div>
                                  </div>
                                ) : null}
                              </div>
                            );
                          })()}
                        </div>
                      </section>
                    )}

                  </div>
                </div>
              )}
            </div>
            </motion.div>
          </div>
        )}
        {viewingPdf && (
          <PdfViewer
            url={viewingPdf.url}
            fileName={viewingPdf.fileName}
            onDownload={() => downloadFileFromUrl(viewingPdf.url, viewingPdf.fileName)}
            onClose={() => {
              if (viewingPdf.url.startsWith("blob:")) URL.revokeObjectURL(viewingPdf.url);
              setViewingPdf(null);
            }}
          />
        )}
      </AnimatePresence>
      <ConvertPrToPoModal
        target={convertPoModal}
        loading={loading}
        onClose={() => !loading && setConvertPoModal(null)}
        onConfirm={handleConvertToPOConfirm}
      />
    </div>
  );
};

// --- Dashboard Component ---

const Dashboard = ({ 
  requests, 
  workflows, 
  user, 
  onStartRequest,
  onViewRequest,
  onViewAllRequests
}: { 
  requests: WorkflowRequest[]; 
  workflows: Workflow[]; 
  user: User;
  onStartRequest: (w: Workflow) => void;
  onViewRequest: (r: WorkflowRequest) => void;
  onViewAllRequests: () => void;
}) => {
  // Calculate stats
  const pendingForMe = requests.filter(r => {
    if (!isWorkflowRequestPending(r)) return false;
    const currentStep = r.template_steps[r.current_step_index];
    if (!currentStep) return false;
    return userCanApproveWorkflowStep(user, r, currentStep);
  });

  const myRequests = requests.filter(r => r.requester_id === user.id);
  
  const myApprovals = requests.filter(r => 
    r.approvals?.some(a => a.approver_id === user.id)
  );

  const stats = [
    { label: 'Pending My Approval', value: pendingForMe.length, icon: Clock, color: 'text-amber-600', bg: 'bg-amber-50' },
    { label: 'My Active Requests', value: myRequests.filter(r => r.status === 'Pending').length, icon: Send, color: 'text-indigo-600', bg: 'bg-indigo-50' },
    { label: 'Total Approved', value: myApprovals.filter(r => r.approvals?.some(a => a.approver_id === user.id && a.status === 'Approved')).length, icon: CheckCircle2, color: 'text-emerald-600', bg: 'bg-emerald-50' },
    { label: 'Total Rejected', value: myApprovals.filter(r => r.approvals?.some(a => a.approver_id === user.id && a.status === 'Rejected')).length, icon: XCircle, color: 'text-red-600', bg: 'bg-red-50' },
  ];

  // Chart data: Requests by category
  const categoryData = workflows.map(w => ({
    name: w.category,
    count: requests.filter(r => r.template_id === w.id).length
  })).reduce((acc: any[], curr) => {
    const existing = acc.find(a => a.name === curr.name);
    if (existing) {
      existing.count += curr.count;
    } else {
      acc.push(curr);
    }
    return acc;
  }, []);

  // Chart data: Status distribution
  const statusData = [
    { name: 'Pending', value: requests.filter(r => r.status === 'Pending').length, color: '#f59e0b' },
    { name: 'Approved', value: requests.filter(r => r.status === 'Approved').length, color: '#10b981' },
    { name: 'Rejected', value: requests.filter(r => r.status === 'Rejected').length, color: '#ef4444' },
  ];

  const recentActivities = requests
    .flatMap(r => (r.approvals || []).map(a => ({ ...a, requestTitle: r.title, requestId: r.id, fullRequest: r })))
    .filter(a => a.approver_id === user.id)
    .sort((a, b) => new Date(b.created_at).getTime() - new Date(a.created_at).getTime())
    .slice(0, 5);

  return (
    <div className="space-y-8">
      {/* Stats Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        {stats.map((stat, i) => (
          <motion.div
            key={stat.label}
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: i * 0.1 }}
            className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm"
          >
            <div className="flex items-center gap-4">
              <div className={cn("w-12 h-12 rounded-xl flex items-center justify-center", stat.bg)}>
                <stat.icon className={cn("w-6 h-6", stat.color)} />
              </div>
              <div>
                <p className="text-xs font-bold text-zinc-400 uppercase tracking-wider">{stat.label}</p>
                <p className="text-2xl font-black text-zinc-900">{stat.value}</p>
              </div>
            </div>
          </motion.div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        {/* Charts Section */}
        <div className="lg:col-span-2 space-y-8">
          <div className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm">
            <h3 className="text-sm font-bold text-zinc-900 uppercase tracking-wider mb-6 flex items-center gap-2">
              <TrendingUp className="w-4 h-4 text-indigo-600" />
              Request Volume by Category
            </h3>
            <div className="h-[300px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={categoryData}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
                  <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#71717a' }} />
                  <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#71717a' }} />
                  <Tooltip 
                    contentStyle={{ backgroundColor: '#fff', borderRadius: '12px', border: '1px solid #e4e4e7', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                    cursor={{ fill: '#f8fafc' }}
                  />
                  <Bar dataKey="count" fill="#4f46e5" radius={[4, 4, 0, 0]} barSize={40} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm">
            <div className="flex items-center justify-between mb-6">
              <h3 className="text-sm font-bold text-zinc-900 uppercase tracking-wider flex items-center gap-2">
                <Clock className="w-4 h-4 text-amber-600" />
                Recent Pending Approvals
              </h3>
              <button 
                onClick={onViewAllRequests}
                className="text-[10px] font-black text-indigo-600 uppercase tracking-widest hover:underline"
              >
                View All
              </button>
            </div>
            <div className="space-y-3">
              {pendingForMe.length === 0 ? (
                <div className="text-center py-8">
                  <CheckCircle2 className="w-10 h-10 text-emerald-200 mx-auto mb-2" />
                  <p className="text-sm text-zinc-400 font-medium">All caught up! No pending approvals.</p>
                </div>
              ) : (
                pendingForMe.slice(0, 5).map(r => (
                  <div 
                    key={r.id} 
                    onClick={() => onViewRequest(r)}
                    className="flex items-center justify-between p-4 rounded-xl border border-zinc-100 hover:bg-zinc-50 transition-all cursor-pointer group"
                  >
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-lg bg-indigo-50 flex items-center justify-center">
                        <FileText className="w-5 h-5 text-indigo-600" />
                      </div>
                      <div>
                        <p className="text-sm font-bold text-zinc-900 group-hover:text-indigo-600 transition-colors">{r.title}</p>
                        <p className="text-[10px] text-zinc-400 uppercase font-bold">{r.requester_name} • {r.template_name}</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] font-bold px-2 py-1 bg-amber-50 text-amber-600 rounded-full border border-amber-100">
                        Step {r.current_step_index + 1}
                      </span>
                      <ChevronRight className="w-4 h-4 text-zinc-300 group-hover:text-indigo-400 transition-colors" />
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>

        {/* Side Section */}
        <div className="space-y-8">
          <div className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm">
            <h3 className="text-sm font-bold text-zinc-900 uppercase tracking-wider mb-6 flex items-center gap-2">
              <Activity className="w-4 h-4 text-indigo-600" />
              Overall Status
            </h3>
            <div className="h-[200px]">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={statusData}
                    cx="50%"
                    cy="50%"
                    innerRadius={60}
                    outerRadius={80}
                    paddingAngle={5}
                    dataKey="value"
                  >
                    {statusData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="flex justify-center gap-4 mt-4">
              {statusData.map(s => (
                <div key={s.name} className="flex items-center gap-1.5">
                  <div className="w-2 h-2 rounded-full" style={{ backgroundColor: s.color }} />
                  <span className="text-[10px] font-bold text-zinc-500 uppercase">{s.name}</span>
                </div>
              ))}
            </div>
          </div>

          <div className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm">
            <h3 className="text-sm font-bold text-zinc-900 uppercase tracking-wider mb-6 flex items-center gap-2">
              <RotateCcw className="w-4 h-4 text-emerald-600" />
              My Recent Actions
            </h3>
            <div className="space-y-4">
              {recentActivities.length === 0 ? (
                <p className="text-xs text-zinc-400 italic text-center py-4">No recent actions found.</p>
              ) : (
                recentActivities.map((activity) => (
                  <div 
                    key={activity.id} 
                    onClick={() => onViewRequest(activity.fullRequest)}
                    className="flex items-start gap-3 cursor-pointer group"
                  >
                    <div className={cn(
                      "mt-1 w-2 h-2 rounded-full shrink-0",
                      activity.status === 'Approved' ? "bg-emerald-500" : "bg-red-500"
                    )} />
                    <div>
                      <p className="text-xs font-bold text-zinc-900 group-hover:text-indigo-600 transition-colors line-clamp-1">{activity.requestTitle}</p>
                      <p className="text-[10px] text-zinc-400">
                        {activity.status === 'Approved' ? 'Approved' : 'Rejected'} on {formatDateMYT(activity.created_at)}
                      </p>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>

          <div className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm">
            <h3 className="text-sm font-bold text-zinc-900 uppercase tracking-wider mb-4 flex items-center gap-2">
              <PlusCircle className="w-4 h-4 text-indigo-600" />
              Quick Start
            </h3>
            <div className="space-y-2">
              {workflows
                .filter(
                  w =>
                    (w.status ?? '').toString().trim().toLowerCase() === 'approved' &&
                    (w.is_active ?? true)
                )
                .slice(0, 3)
                .map(w => (
                <button
                  key={w.id}
                  onClick={() => onStartRequest(w)}
                  className="w-full flex items-center justify-between p-3 rounded-xl border border-zinc-100 hover:border-indigo-200 hover:bg-indigo-50 transition-all group text-left"
                >
                  <span className="text-xs font-bold text-zinc-600 group-hover:text-indigo-600">{w.name}</span>
                  <Plus className="w-4 h-4 text-zinc-300 group-hover:text-indigo-400" />
                </button>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

const WorkflowList = ({ workflows, user, roles: availableRoles, onRefresh, onStartRequest }: { workflows: Workflow[], user: User, roles: Role[], onRefresh: () => void, onStartRequest: (w: Workflow) => void }) => {
  const [selectedWorkflow, setSelectedWorkflow] = useState<Workflow | null>(null);
  const [loading, setLoading] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [filter, setFilter] = useState<'all' | 'Pending' | 'Approved' | 'Rejected'>('all');
  const [editData, setEditData] = useState({ 
    name: '', 
    category: 'general' as string,
    steps: [] as WorkflowStep[], 
    table_columns: [] as string[],
    attachments_required: false
  });
  const [editNewColumnName, setEditNewColumnName] = useState('');
  const [users, setUsers] = useState<User[]>([]);
  const isPR = isPRTemplate(editData);
  const isPO = isPOTemplate(editData);
  const isSR = isSRTemplate(editData);

  useEffect(() => {
    if (isPR) setEditData(prev => ({ ...prev, steps: FIXED_PR_STEPS }));
    else if (isPO) setEditData(prev => ({ ...prev, steps: FIXED_PO_STEPS_FULL }));
    else if (isSR) setEditData(prev => ({ ...prev, steps: FIXED_SR_STEPS }));
  }, [isPR, isPO, isSR]);

  const filteredWorkflows = workflows.filter(w => filter === 'all' || w.status === filter);
  
  useEffect(() => {
    const fetchUsers = async () => {
      try {
        const data = await api.request('/api/users');
        setUsers(data);
      } catch (err) {
        console.error('Failed to fetch users', err);
      }
    };
    fetchUsers();
  }, []);

  const hasPermission = (permission: string) => {
    if (!user) return false;
    if (user.permissions?.includes('admin')) return true;
    return user.permissions?.includes(permission);
  };

  const isAdmin = user.roles?.some(r => r.toLowerCase() === 'admin') || hasPermission('admin');
  const isDirector = user.roles?.some(r => r.toLowerCase() === 'director') && (user.department || '').toLowerCase() === 'management';
  const canApproveTemplates = hasPermission('approve_templates');
  const canCreateTemplates = hasPermission('create_templates');

  const handleStatusUpdate = async (id: number, status: string) => {
    try {
      await api.request(`/api/workflows/${id}/status`, {
        method: 'PATCH',
        body: JSON.stringify({ status }),
      });
      toast.success(`Workflow Template ${status}`);
      onRefresh();
      setSelectedWorkflow(null);
    } catch (err: any) {
      toast.error(err.message);
    }
  };

  const handleActiveToggle = async (id: number, isActive: boolean) => {
    try {
      await api.request(`/api/workflows/${id}/active`, {
        method: 'PATCH',
        body: JSON.stringify({ is_active: isActive }),
      });
      toast.success(isActive ? 'Template activated' : 'Template deactivated');
      onRefresh();
      if (selectedWorkflow && selectedWorkflow.id === id) {
        setSelectedWorkflow({ ...selectedWorkflow, is_active: isActive });
      }
    } catch (err: any) {
      toast.error(err.message);
    }
  };

  const handleSaveEdit = async () => {
    if (!selectedWorkflow) return;
    setLoading(true);
    try {
      await api.request(`/api/workflows/${selectedWorkflow.id}`, {
        method: 'PATCH',
        body: JSON.stringify(editData),
      });
      toast.success('Template updated');
      setIsEditing(false);
      onRefresh();
      setSelectedWorkflow({ ...selectedWorkflow, ...editData });
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const addEditColumn = () => {
    const name = editNewColumnName.trim();
    if (!name) return;
    if (editData.table_columns.includes(name)) {
      setEditNewColumnName('');
      return;
    }
    setEditData({ ...editData, table_columns: [...editData.table_columns, name] });
    setEditNewColumnName('');
  };

  const removeEditColumn = (idx: number) => {
    setEditData({ ...editData, table_columns: editData.table_columns.filter((_, i) => i !== idx) });
  };

  const updateEditColumn = (idx: number, nextName: string) => {
    setEditData({
      ...editData,
      table_columns: editData.table_columns.map((c, i) => (i === idx ? nextName : c)),
    });
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between mb-2">
        <h2 className="text-xl font-bold text-zinc-900">
          {(isAdmin || isDirector || canApproveTemplates || canCreateTemplates) ? 'All Workflow Templates' : 'Shared Workflow Templates'}
        </h2>
        <span className="text-xs text-zinc-500 bg-zinc-100 px-2 py-1 rounded-full">
        {filteredWorkflows.length} Templates
          {/* {workflows.length} Templates */}
        </span>
      </div>
      
      <div className="flex gap-2 mb-4">
        {['all', 'pending', 'approved', 'rejected'].map((f) => (
          <button
            key={f}
            onClick={() => setFilter(f as any)}
            className={cn(
              "px-3 py-1.5 rounded-lg text-xs font-bold capitalize transition-all",
              filter === f ? "bg-indigo-600 text-white" : "bg-white text-zinc-500 border border-zinc-200 hover:border-indigo-200"
            )}
          >
            {f}
          </button>
        ))}
      </div>

      <div className="grid grid-cols-1 gap-4">
        {filteredWorkflows.length === 0 && (
          <div className="text-center py-12 bg-white rounded-xl border border-dashed border-zinc-300">
            <ClipboardList className="w-12 h-12 text-zinc-300 mx-auto mb-3" />
            <p className="text-zinc-500">No {filter !== 'all' ? filter : ''} templates found.</p>
            {/* <p className="text-zinc-500">No templates found.</p> */}
          </div>
        )}
        {filteredWorkflows.map((w) => (
          <div
            key={w.id}
            onClick={() => {
              setSelectedWorkflow(w);
            }}
            className="bg-white p-5 rounded-xl border border-zinc-200 hover:border-indigo-300 hover:shadow-md transition-all cursor-pointer group"
          >
            <div className="flex items-start justify-between">
              <div>
                <div className="flex items-center gap-2 mb-1">
                  <h3 className="font-bold text-zinc-900 group-hover:text-indigo-600 transition-colors">{w.name}</h3>
                  <span className={cn(
                    "text-[10px] uppercase font-bold px-2 py-0.5 rounded-full",
                    (w.status ?? '').toString().trim().toLowerCase() === 'approved' ? "bg-emerald-100 text-emerald-700" :
                    (w.status ?? '').toString().trim().toLowerCase() === 'rejected' ? "bg-red-100 text-red-700" :
                    "bg-amber-100 text-amber-700"
                  )}>
                    {w.status}
                  </span>
                  <span className="text-[10px] uppercase font-bold px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-600 border border-indigo-100">
                    {w.category}
                  </span>
                  <span className={cn(
                    "text-[10px] uppercase font-bold px-2 py-0.5 rounded-full border",
                    (w.is_active ?? true)
                      ? "bg-emerald-50 text-emerald-700 border-emerald-100"
                      : "bg-zinc-100 text-zinc-500 border-zinc-200"
                  )}>
                    {(w.is_active ?? true) ? 'Active' : 'Inactive'}
                  </span>
                </div>
                <div className="flex items-center gap-4 mt-3 text-xs text-zinc-400">
                  <span className="flex items-center gap-1">
                    <Clock className="w-3 h-3" />
                    {formatDateMYT(w.created_at)}
                  </span>
                  {isAdmin && (
                    <span className="flex items-center gap-1">
                      <Users className="w-3 h-3" />
                      By {w.creator_name}
                    </span>
                  )}
                  <span className="flex items-center gap-1">
                    <ChevronRight className="w-3 h-3" />
                    {w.steps.length} Steps
                  </span>
                </div>
              </div>
              <div className="flex items-center gap-2">
                {(w.status ?? '').toString().trim().toLowerCase() === 'approved' && (w.is_active ?? true) && (
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      onStartRequest(w);
                    }}
                    className="bg-indigo-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold hover:bg-indigo-700 transition-colors flex items-center gap-1"
                  >
                    <Send className="w-3 h-3" />
                    Start Request
                  </button>
                )}
                <div className="w-8 h-8 rounded-full bg-zinc-50 flex items-center justify-center group-hover:bg-indigo-50 transition-colors">
                  <ChevronRight className="w-4 h-4 text-zinc-400 group-hover:text-indigo-600" />
                </div>
              </div>
            </div>
          </div>
        ))}
      </div>

      <AnimatePresence>
        {selectedWorkflow && (
          <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white w-full max-w-2xl rounded-2xl shadow-2xl overflow-hidden"
            >
              <div className="p-6 border-b border-zinc-100 flex items-center justify-between bg-zinc-50">
                <div className="flex items-center gap-3">
                  <div>
                    <h2 className="text-xl font-bold text-zinc-900">{selectedWorkflow.name}</h2>
                    <p className="text-sm text-zinc-500">Template created on {formatDateTimeMYT(selectedWorkflow.created_at)}</p>
                  </div>
                  <span className={cn(
                    "text-[10px] uppercase font-bold px-2 py-0.5 rounded-full",
                    selectedWorkflow.status === 'approved' ? "bg-emerald-100 text-emerald-700" :
                    selectedWorkflow.status === 'rejected' ? "bg-red-100 text-red-700" :
                    "bg-amber-100 text-amber-700"
                  )}>
                    {selectedWorkflow.status}
                  </span>
                  <span className={cn(
                    "text-[10px] uppercase font-bold px-2 py-0.5 rounded-full border",
                    (selectedWorkflow.is_active ?? true)
                      ? "bg-emerald-50 text-emerald-700 border-emerald-100"
                      : "bg-zinc-100 text-zinc-500 border-zinc-200"
                  )}>
                    {(selectedWorkflow.is_active ?? true) ? 'Active' : 'Inactive'}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  {(canCreateTemplates || canApproveTemplates || isAdmin) && !isEditing && (
                    <button
                      onClick={() => handleActiveToggle(selectedWorkflow.id, !(selectedWorkflow.is_active ?? true))}
                      className={cn(
                        "flex items-center gap-2 px-3 py-1.5 rounded-lg border text-xs font-bold transition-all",
                        (selectedWorkflow.is_active ?? true)
                          ? "border-zinc-200 text-zinc-600 hover:bg-zinc-100"
                          : "border-emerald-200 text-emerald-700 hover:bg-emerald-50"
                      )}
                    >
                      {(selectedWorkflow.is_active ?? true) ? 'Deactivate' : 'Activate'}
                    </button>
                  )}
                  {(canCreateTemplates || (canApproveTemplates && selectedWorkflow.status === 'pending')) && !isEditing && (
                    <button 
                      onClick={() => {
                        setEditData({ 
                          name: selectedWorkflow.name, 
                          category: (selectedWorkflow as Workflow).category || 'general',
                          steps: [...selectedWorkflow.steps],
                          table_columns: [...(selectedWorkflow.table_columns || [])],
                          attachments_required: !!selectedWorkflow.attachments_required
                        });
                        setEditNewColumnName('');
                        setIsEditing(true);
                      }}
                      className="flex items-center gap-2 px-3 py-1.5 rounded-lg border border-indigo-200 text-indigo-600 text-xs font-bold hover:bg-indigo-50 transition-all"
                    >
                      <Edit2 className="w-3.5 h-3.5" />
                      Edit Template
                    </button>
                  )}
                  <button onClick={() => { setSelectedWorkflow(null); setIsEditing(false); }} className="text-zinc-400 hover:text-zinc-600">
                    <XCircle className="w-6 h-6" />
                  </button>
                </div>
              </div>
              <div className="p-6 space-y-6 max-h-[70vh] overflow-y-auto">
                {isEditing ? (
                  <div className="space-y-4">
                    <div>
                      <label className="text-xs font-bold text-zinc-500 uppercase mb-1 block">Template Name</label>
                      <input
                        type="text"
                        className="w-full px-3 py-2 rounded-lg border border-zinc-300 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                        value={editData.name}
                        onChange={(e) => setEditData({ ...editData, name: e.target.value })}
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-xs font-bold text-zinc-500 uppercase mb-1 block">Data Table Columns</label>
                      <div className="flex items-center justify-between gap-3">
                        <input
                          type="text"
                          placeholder="e.g., Amount, Qty"
                          className="flex-1 text-xs px-3 py-2 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
                          value={editNewColumnName}
                          onChange={(e) => setEditNewColumnName(e.target.value)}
                          onKeyPress={(e) => e.key === 'Enter' && (e.preventDefault(), addEditColumn())}
                        />
                        <button
                          type="button"
                          onClick={addEditColumn}
                          className="text-xs bg-emerald-50 text-emerald-600 px-3 py-2 rounded-lg font-bold hover:bg-emerald-100 transition-colors shrink-0"
                        >
                          + Add Column
                        </button>
                      </div>

                      {editData.table_columns.length === 0 && (
                        <p className="text-xs text-zinc-400 italic">No custom columns defined.</p>
                      )}

                      {editData.table_columns.length > 0 && (
                        <div className="flex flex-wrap gap-2">
                          {editData.table_columns.map((col, idx) => (
                            <div
                              key={`${col}-${idx}`}
                              className="flex items-center gap-2 bg-zinc-100 px-3 py-1 rounded-full text-xs text-zinc-600 border border-zinc-200"
                            >
                              <input
                                type="text"
                                className="bg-transparent border-none focus:ring-0 text-xs w-28 outline-none font-medium"
                                value={col}
                                onChange={(e) => updateEditColumn(idx, e.target.value)}
                              />
                              <button
                                type="button"
                                onClick={() => removeEditColumn(idx)}
                                className="text-zinc-400 hover:text-red-500"
                                title="Remove column"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            </div>
                          ))}
                        </div>
                      )}

                      {(isPR || isPO) && (
                        <p className="text-[11px] text-zinc-400 pt-1">
                          Note: procurement PR/PO templates use fixed columns in request forms.
                        </p>
                      )}
                    </div>

                    <div>
                      <label className="text-xs font-bold text-zinc-500 uppercase mb-2 block">Approval Steps</label>
                      <div className="space-y-2">
                        {editData.steps.map((step, idx) => (
                          <div key={step.id || idx} className="p-3 bg-zinc-50 rounded-lg border border-zinc-200 flex gap-2 items-center">
                            <div className="w-6 h-6 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px] font-bold shrink-0">
                              {idx + 1}
                            </div>
                            <input
                              type="text"
                              className="flex-1 px-2 py-1 rounded border border-zinc-300 text-xs outline-none focus:ring-1 focus:ring-indigo-500"
                              disabled={isPR || isPO}
                              value={step.label}
                              onChange={(e) => {
                                const newSteps = [...editData.steps];
                                newSteps[idx] = { ...newSteps[idx], label: e.target.value };
                                setEditData({ ...editData, steps: newSteps });
                              }}
                              placeholder="Step Label"
                            />
                            <select
                              className="px-2 py-1 rounded border border-zinc-300 text-xs outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
                              disabled={isPR || isPO}
                              value={step.approverRole}
                              onChange={(e) => {
                                const newSteps = [...editData.steps];
                                newSteps[idx] = { ...newSteps[idx], approverRole: e.target.value };
                                setEditData({ ...editData, steps: newSteps });
                              }}
                            >
                              <option value="Preparer">Preparer</option>
                              <option value="Checker">Checker</option>
                              <option value="Approver">Approver</option>
                              <option value="Director">Director</option>
                              {availableRoles.filter(r => !['admin', 'user', 'preparer', 'checker', 'approver', 'director'].includes(r.name.toLowerCase())).map(r => (
                                <option key={r.id} value={r.name}>{r.name}</option>
                              ))}
                            </select>
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="flex items-center gap-2 bg-zinc-50 p-3 rounded-lg border border-zinc-200">
                      <input
                        type="checkbox"
                        id="edit_attachments_required"
                        className="w-4 h-4 text-indigo-600 rounded border-zinc-300 focus:ring-indigo-500"
                        checked={editData.attachments_required}
                        onChange={(e) => setEditData({ ...editData, attachments_required: e.target.checked })}
                      />
                      <label htmlFor="edit_attachments_required" className="text-sm font-medium text-zinc-700 cursor-pointer">
                        Require attachments when submitting requests
                      </label>
                    </div>
                    <div className="flex gap-3 pt-4">
                      <button
                        onClick={() => setIsEditing(false)}
                        className="flex-1 px-4 py-2 rounded-lg border border-zinc-200 text-zinc-600 font-bold hover:bg-zinc-50 transition-all"
                      >
                        Cancel
                      </button>
                      <button
                        onClick={handleSaveEdit}
                        disabled={loading}
                        className="flex-1 px-4 py-2 rounded-lg bg-indigo-600 text-white font-bold hover:bg-indigo-700 transition-all disabled:opacity-50"
                      >
                        {loading ? 'Saving...' : 'Save Changes'}
                      </button>
                    </div>
                  </div>
                ) : (
                  <>
                {selectedWorkflow.table_columns && selectedWorkflow.table_columns.length > 0 && (
                  <section>
                    <h4 className="text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Data Table Columns</h4>
                    <div className="flex flex-wrap gap-2">
                      {selectedWorkflow.table_columns.map((col, i) => (
                        <span key={`${col}-${i}`} className="bg-zinc-100 px-3 py-1 rounded-full text-xs text-zinc-600 border border-zinc-200">
                          {col}
                        </span>
                      ))}
                    </div>
                  </section>
                )}

                <section>
                  <h4 className="text-xs font-bold text-zinc-400 uppercase tracking-wider mb-3">Approval Chain</h4>
                  <div className="space-y-3">
                    {selectedWorkflow.steps.map((step, i) => (
                      <div key={step.id} className="flex items-center gap-4">
                        <div className="w-8 h-8 rounded-full bg-indigo-600 text-white flex items-center justify-center text-xs font-bold shrink-0">
                          {i + 1}
                        </div>
                        <div className="flex-1 bg-zinc-50 p-3 rounded-lg border border-zinc-200">
                          <p className="font-semibold text-zinc-900 text-sm">{step.label}</p>
                          <p className="text-xs text-zinc-500">Approver: {step.approverRole}</p>
                        </div>
                      </div>
                    ))}
                  </div>
                </section>

                <section>
                  <h4 className="text-xs font-bold text-zinc-400 uppercase tracking-wider mb-2">Requirements</h4>
                  <div className="flex items-center gap-2 text-sm text-zinc-600">
                    <div className={cn(
                      "w-4 h-4 rounded-full flex items-center justify-center",
                      selectedWorkflow.attachments_required ? "bg-indigo-100 text-indigo-600" : "bg-zinc-100 text-zinc-400"
                    )}>
                      {selectedWorkflow.attachments_required ? <CheckCircle2 className="w-3 h-3" /> : <XCircle className="w-3 h-3" />}
                    </div>
                    {selectedWorkflow.attachments_required ? "Attachments are required for requests" : "Attachments are optional"}
                  </div>
                </section>
              </>
            )}
          </div>
              <div className="p-6 bg-zinc-50 border-t border-zinc-100 flex gap-3">
                {(isAdmin || canApproveTemplates) && selectedWorkflow.status === 'pending' && (
                  <>
                    <button
                      onClick={() => handleStatusUpdate(selectedWorkflow.id, 'approved')}
                      className="flex-1 bg-emerald-600 text-white py-2 rounded-lg font-bold hover:bg-emerald-700 transition-colors flex items-center justify-center gap-2"
                    >
                      <CheckCircle2 className="w-4 h-4" />
                      Approve Template
                    </button>
                    <button
                      onClick={() => handleStatusUpdate(selectedWorkflow.id, 'rejected')}
                      className="flex-1 bg-red-600 text-white py-2 rounded-lg font-bold hover:bg-red-700 transition-colors flex items-center justify-center gap-2"
                    >
                      <XCircle className="w-4 h-4" />
                      Reject
                    </button>
                  </>
                )}
                {selectedWorkflow.status === 'approved' && (selectedWorkflow.is_active ?? true) && (
                  <button
                    onClick={() => {
                      onStartRequest(selectedWorkflow);
                      setSelectedWorkflow(null);
                    }}
                    className="flex-1 bg-indigo-600 text-white py-2 rounded-lg font-bold hover:bg-indigo-700 transition-colors flex items-center justify-center gap-2"
                  >
                    <Send className="w-4 h-4" />
                    Start Request from this Template
                  </button>
                )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

/** Purchasing (and admins) maintain the entity-scoped cost center catalog (`dbo.cost_center`). */
function CostCentersAdminPage({ entity }: { entity: string | null }) {
  const ent = String(entity || '').trim();
  const [rows, setRows] = useState<
    Array<{
      id: number;
      entity: string;
      code: string;
      name: string;
      gl_account: string | null;
      status: boolean;
      created_at?: string;
    }>
  >([]);
  const [loading, setLoading] = useState(false);
  const [showInactive, setShowInactive] = useState(true);
  const [newRow, setNewRow] = useState({
    code: '',
    name: '',
    gl_account: '',
    active: true,
  });
  const [editing, setEditing] = useState<{
    id: number;
    code: string;
    name: string;
    gl_account: string;
    active: boolean;
  } | null>(null);

  const load = useCallback(async () => {
    if (!ent) {
      setRows([]);
      return;
    }
    setLoading(true);
    try {
      const q = showInactive ? '?include_inactive=1' : '';
      const data = await api.request(`/api/cost-centers${q}`);
      const list = Array.isArray(data) ? data : [];
      setRows(
        list.map((r: any) => ({
          id: Number(r.id),
          entity: String(r.entity || ''),
          code: String(r.code || ''),
          name: String(r.name || ''),
          gl_account: r.gl_account != null ? String(r.gl_account) : null,
          status: !!r.status,
          created_at: r.created_at,
        }))
      );
    } catch (e: unknown) {
      toast.error(e instanceof Error ? e.message : 'Failed to load cost centers');
      setRows([]);
    } finally {
      setLoading(false);
    }
  }, [ent, showInactive]);

  useEffect(() => {
    void load();
  }, [load]);

  const handleCreate = async () => {
    const code = newRow.code.trim();
    const name = newRow.name.trim();
    if (!code || !name) {
      toast.error('Code and name are required');
      return;
    }
    try {
      await api.request('/api/cost-centers', {
        method: 'POST',
        body: JSON.stringify({
          code,
          name,
          gl_account: newRow.gl_account.trim() || undefined,
          status: newRow.active ? 1 : 0,
        }),
      });
      toast.success('Cost center added');
      setNewRow({ code: '', name: '', gl_account: '', active: true });
      void load();
    } catch (e: unknown) {
      toast.error(e instanceof Error ? e.message : 'Failed to create');
    }
  };

  const saveEdit = async () => {
    if (!editing) return;
    const code = editing.code.trim();
    const name = editing.name.trim();
    if (!code || !name) {
      toast.error('Code and name are required');
      return;
    }
    try {
      await api.request(`/api/cost-centers/${editing.id}`, {
        method: 'PATCH',
        body: JSON.stringify({
          code,
          name,
          gl_account: editing.gl_account.trim() || null,
          status: editing.active ? 1 : 0,
        }),
      });
      toast.success('Saved');
      setEditing(null);
      void load();
    } catch (e: unknown) {
      toast.error(e instanceof Error ? e.message : 'Failed to save');
    }
  };

  if (!ent) {
    return (
      <div className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-950">
        Select an <strong>entity</strong> in the header to load and edit the cost center list for that entity.
      </div>
    );
  }

  return (
    <div className="flex flex-col flex-1 min-h-0 gap-4">
      <div className="flex flex-wrap items-end justify-between gap-3 shrink-0">
        <div>
          <h1 className="text-xl font-bold text-zinc-900">Cost centers</h1>
          <p className="text-xs text-zinc-500 mt-1">
            Entity <span className="font-semibold text-zinc-700">{ent}</span> — codes must be unique per entity. Inactive rows stay valid for existing PR/PO lines but are hidden from new picks.
          </p>
        </div>
        <label className="flex items-center gap-2 text-sm text-zinc-600 cursor-pointer select-none">
          <input
            type="checkbox"
            className="rounded border-zinc-300"
            checked={showInactive}
            onChange={(e) => setShowInactive(e.target.checked)}
          />
          Show inactive
        </label>
      </div>

      <div className="rounded-xl border border-zinc-200 bg-white p-4 shadow-sm shrink-0">
        <h2 className="text-sm font-bold text-zinc-900 mb-3">Add cost center</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3 items-end">
          <div className="lg:col-span-1">
            <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider block mb-1">Code</label>
            <input
              value={newRow.code}
              onChange={(e) => setNewRow({ ...newRow, code: e.target.value })}
              className="w-full rounded-lg border border-zinc-200 px-2.5 py-2 text-sm"
              placeholder="e.g. 1000"
            />
          </div>
          <div className="lg:col-span-2">
            <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider block mb-1">Name</label>
            <input
              value={newRow.name}
              onChange={(e) => setNewRow({ ...newRow, name: e.target.value })}
              className="w-full rounded-lg border border-zinc-200 px-2.5 py-2 text-sm"
              placeholder="Department / description"
            />
          </div>
          <div>
            <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider block mb-1">GL account</label>
            <input
              value={newRow.gl_account}
              onChange={(e) => setNewRow({ ...newRow, gl_account: e.target.value })}
              className="w-full rounded-lg border border-zinc-200 px-2.5 py-2 text-sm"
              placeholder="Optional"
            />
          </div>
          <div className="flex items-center gap-3">
            <label className="flex items-center gap-2 text-sm text-zinc-600 cursor-pointer">
              <input
                type="checkbox"
                className="rounded border-zinc-300"
                checked={newRow.active}
                onChange={(e) => setNewRow({ ...newRow, active: e.target.checked })}
              />
              Active
            </label>
            <button
              type="button"
              onClick={() => void handleCreate()}
              className="rounded-lg bg-indigo-600 text-white px-4 py-2 text-sm font-semibold hover:bg-indigo-700 shrink-0"
            >
              Add
            </button>
          </div>
        </div>
      </div>

      <div className="flex-1 min-h-0 overflow-auto rounded-xl border border-zinc-200 bg-white flex flex-col">
        <div className="px-4 py-2 border-b border-zinc-100 flex items-center justify-between shrink-0">
          <span className="text-xs font-bold text-zinc-500 uppercase tracking-wider">Catalog</span>
          <button
            type="button"
            onClick={() => void load()}
            className="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-600 hover:text-indigo-800"
          >
            <RefreshCw className="w-3.5 h-3.5" />
            Refresh
          </button>
        </div>
        {loading ? (
          <div className="p-8 text-center text-sm text-zinc-500">Loading…</div>
        ) : rows.length === 0 ? (
          <div className="p-8 text-center text-sm text-zinc-500">No cost centers for this entity yet.</div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b border-zinc-100 text-left text-[10px] font-bold uppercase tracking-wider text-zinc-400">
                  <th className="px-4 py-2">Code</th>
                  <th className="px-4 py-2">Name</th>
                  <th className="px-4 py-2">GL</th>
                  <th className="px-4 py-2">Active</th>
                  <th className="px-4 py-2 w-36">Actions</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row) => (
                  <tr key={row.id} className="border-b border-zinc-50 hover:bg-zinc-50/80">
                    {editing?.id === row.id ? (
                      <>
                        <td className="px-4 py-2 align-top">
                          <input
                            value={editing.code}
                            onChange={(e) => setEditing({ ...editing, code: e.target.value })}
                            className="w-full min-w-[4rem] rounded border border-zinc-200 px-2 py-1 text-sm"
                          />
                        </td>
                        <td className="px-4 py-2 align-top">
                          <input
                            value={editing.name}
                            onChange={(e) => setEditing({ ...editing, name: e.target.value })}
                            className="w-full min-w-[8rem] rounded border border-zinc-200 px-2 py-1 text-sm"
                          />
                        </td>
                        <td className="px-4 py-2 align-top">
                          <input
                            value={editing.gl_account}
                            onChange={(e) => setEditing({ ...editing, gl_account: e.target.value })}
                            className="w-full min-w-[5rem] rounded border border-zinc-200 px-2 py-1 text-sm"
                          />
                        </td>
                        <td className="px-4 py-2 align-top">
                          <input
                            type="checkbox"
                            className="rounded border-zinc-300"
                            checked={editing.active}
                            onChange={(e) => setEditing({ ...editing, active: e.target.checked })}
                          />
                        </td>
                        <td className="px-4 py-2 align-top whitespace-nowrap">
                          <button
                            type="button"
                            onClick={() => void saveEdit()}
                            className="text-xs font-semibold text-emerald-600 hover:text-emerald-800 mr-3"
                          >
                            Save
                          </button>
                          <button
                            type="button"
                            onClick={() => setEditing(null)}
                            className="text-xs font-semibold text-zinc-500 hover:text-zinc-800"
                          >
                            Cancel
                          </button>
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="px-4 py-2 font-mono text-xs">{row.code}</td>
                        <td className="px-4 py-2">{row.name}</td>
                        <td className="px-4 py-2 text-zinc-600">{row.gl_account || '—'}</td>
                        <td className="px-4 py-2">{row.status ? 'Yes' : 'No'}</td>
                        <td className="px-4 py-2">
                          <button
                            type="button"
                            onClick={() =>
                              setEditing({
                                id: row.id,
                                code: row.code,
                                name: row.name,
                                gl_account: row.gl_account ?? '',
                                active: row.status,
                              })
                            }
                            className="text-xs font-semibold text-indigo-600 hover:text-indigo-800 inline-flex items-center gap-1"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                            Edit
                          </button>
                        </td>
                      </>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

const AVAILABLE_PERMISSIONS = [
  { key: 'admin', label: 'Full Admin Access', description: 'Can perform all actions in the system' },
  { key: 'manage_users', label: 'Manage Users', description: 'Can create and edit users' },
  { key: 'manage_cost_centers', label: 'Manage cost centers', description: 'Create and edit the cost center catalog for the active entity (purchasing)' },
  { key: 'create_templates', label: 'Create Templates', description: 'Can create and manage workflow templates' },
  { key: 'approve_templates', label: 'Approve Templates', description: 'Can approve or reject new workflow templates' },
  { key: 'view_history', label: 'View All History', description: 'Can view all workflow requests across departments' },
  { key: 'edit_requests', label: 'Edit Requests', description: 'Can edit pending requests during approval' },
];

const RoleManager = ({ roles, onRefresh }: { roles: Role[], onRefresh: () => void }) => {
  const [newRoleName, setNewRoleName] = useState('');
  const [newRolePermissions, setNewRolePermissions] = useState<string[]>([]);
  const [editingRole, setEditingRole] = useState<Role | null>(null);
  const [loading, setLoading] = useState(false);

  const handleCreate = async () => {
    if (!newRoleName.trim()) return;
    setLoading(true);
    try {
      await api.request('/api/roles', {
        method: 'POST',
        body: JSON.stringify({ 
          name: newRoleName.trim(),
          permissions: newRolePermissions
        }),
      });
      setNewRoleName('');
      setNewRolePermissions([]);
      onRefresh();
      toast.success('Role created');
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleUpdate = async () => {
    if (!editingRole || !editingRole.name.trim()) return;
    setLoading(true);
    try {
      await api.request(`/api/roles/${editingRole.id}`, {
        method: 'PATCH',
        body: JSON.stringify({ 
          name: editingRole.name.trim(),
          permissions: editingRole.permissions
        }),
      });
      setEditingRole(null);
      onRefresh();
      toast.success('Role updated');
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const togglePermission = (role: 'new' | 'editing', permKey: string) => {
    if (role === 'new') {
      setNewRolePermissions(prev => 
        prev.includes(permKey) ? prev.filter(p => p !== permKey) : [...prev, permKey]
      );
    } else if (editingRole) {
      const currentPerms = editingRole.permissions || [];
      const newPerms = currentPerms.includes(permKey)
        ? currentPerms.filter(p => p !== permKey)
        : [...currentPerms, permKey];
      setEditingRole({ ...editingRole, permissions: newPerms });
    }
  };

  return (
    <div className="bg-white rounded-xl border border-zinc-200 p-6 shadow-sm mb-6">
      <h3 className="text-lg font-bold text-zinc-900 mb-4 flex items-center gap-2">
        <Shield className="w-5 h-5 text-indigo-600" />
        Manage Custom Roles
      </h3>
      
      <div className="space-y-4 mb-6">
        <div className="flex gap-2">
          <input
            type="text"
            placeholder="New role name..."
            className="flex-1 px-4 py-2 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
            value={newRoleName}
            onChange={(e) => setNewRoleName(e.target.value)}
            onKeyPress={(e) => e.key === 'Enter' && handleCreate()}
          />
          <button
            onClick={handleCreate}
            disabled={loading || !newRoleName.trim()}
            className="bg-indigo-600 text-white px-4 py-2 rounded-lg font-bold text-sm hover:bg-indigo-700 transition-colors disabled:opacity-50"
          >
            Add Role
          </button>
        </div>
        
        <div className="bg-zinc-50 p-4 rounded-lg border border-zinc-200">
          <p className="text-[10px] font-bold text-zinc-500 uppercase mb-2">Assign Permissions to New Role</p>
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
            {AVAILABLE_PERMISSIONS.map(perm => (
              <label key={perm.key} className="flex items-start gap-2 p-2 rounded border border-white hover:border-zinc-200 cursor-pointer transition-colors">
                <input
                  type="checkbox"
                  checked={newRolePermissions.includes(perm.key)}
                  onChange={() => togglePermission('new', perm.key)}
                  className="mt-1 w-3 h-3 rounded text-indigo-600 focus:ring-indigo-500"
                />
                <div>
                  <p className="text-xs font-bold text-zinc-700">{perm.label}</p>
                  <p className="text-[9px] text-zinc-500 leading-tight">{perm.description}</p>
                </div>
              </label>
            ))}
          </div>
        </div>
      </div>

      <div className="space-y-2">
        {roles.map((role) => (
          <div key={role.id} className="flex flex-col p-3 bg-zinc-50 rounded-lg border border-zinc-200 group">
            {editingRole?.id === role.id ? (
              <div className="space-y-3">
                <div className="flex gap-2">
                  <input
                    type="text"
                    className="flex-1 px-2 py-1 rounded border border-zinc-300 text-sm outline-none focus:ring-1 focus:ring-indigo-500"
                    value={editingRole.name}
                    onChange={(e) => setEditingRole({ ...editingRole, name: e.target.value })}
                    autoFocus
                  />
                  <button onClick={handleUpdate} className="text-xs font-bold text-emerald-600 hover:underline">Save</button>
                  <button onClick={() => setEditingRole(null)} className="text-xs font-bold text-zinc-400 hover:underline">Cancel</button>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 p-2 bg-white rounded border border-zinc-200">
                  {AVAILABLE_PERMISSIONS.map(perm => (
                    <label key={perm.key} className="flex items-center gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={(editingRole.permissions || []).includes(perm.key)}
                        onChange={() => togglePermission('editing', perm.key)}
                        className="w-3 h-3 rounded text-indigo-600 focus:ring-indigo-500"
                      />
                      <span className="text-[10px] text-zinc-600">{perm.label}</span>
                    </label>
                  ))}
                </div>
              </div>
            ) : (
              <div className="flex items-center justify-between">
                <div className="flex flex-col gap-1">
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-medium text-zinc-700 uppercase tracking-wider">{role.name}</span>
                    {['admin', 'user'].includes(role.name) && (
                      <span className="text-[8px] bg-zinc-200 text-zinc-500 px-1.5 py-0.5 rounded font-bold uppercase">System</span>
                    )}
                  </div>
                  <div className="flex flex-wrap gap-1">
                    {(role.permissions || []).map((p, i) => (
                      <span key={`${p}-${i}`} className="text-[8px] bg-indigo-50 text-indigo-600 px-1.5 py-0.5 rounded border border-indigo-100 font-medium">
                        {AVAILABLE_PERMISSIONS.find(ap => ap.key === p)?.label || p}
                      </span>
                    ))}
                    {(!role.permissions || role.permissions.length === 0) && (
                      <span className="text-[8px] text-zinc-400 italic">No special permissions</span>
                    )}
                  </div>
                </div>
                {!['admin', 'user'].includes(role.name) && (
                  <div className="flex items-center gap-3 opacity-0 group-hover:opacity-100 transition-opacity">
                    <button onClick={() => setEditingRole(role)} className="text-zinc-400 hover:text-indigo-600">
                      <Edit2 className="w-4 h-4" />
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
};

const UserModal = ({ 
  user, 
  availableRoles, 
  onClose, 
  onSuccess 
}: { 
  user?: User | null, 
  availableRoles: Role[], 
  onClose: () => void, 
  onSuccess: () => void 
}) => {
  const [username, setUsername] = useState(user?.username || '');
  const [password, setPassword] = useState('');
  const [department, setDepartment] = useState(user?.department || 'General');
  const [selectedRoles, setSelectedRoles] = useState<string[]>(user?.roles || ['user']);
  const [loading, setLoading] = useState(false);

  const departments = ['IT', 'HR', 'Finance', 'Sales', 'Management', 'General'];

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (selectedRoles.length === 0) {
      toast.error('User must have at least one role');
      return;
    }
    setLoading(true);
    try {
      const endpoint = user ? `/api/users/${user.id}` : '/api/users';
      const method = user ? 'PATCH' : 'POST';
      const body: any = { username, roles: selectedRoles, department };
      if (password) body.password = password;

      await api.request(endpoint, {
        method,
        body: JSON.stringify(body),
      });
      toast.success(user ? 'User updated' : 'User created');
      onSuccess();
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const toggleRole = (role: string) => {
    setSelectedRoles(prev => 
      prev.includes(role) ? prev.filter(r => r !== role) : [...prev, role]
    );
  };

  return (
    <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50 p-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden"
      >
        <div className="p-6 border-b border-zinc-100 flex justify-between items-center">
          <h3 className="text-xl font-bold text-zinc-900">{user ? 'Edit User' : 'Create New User'}</h3>
          <button onClick={onClose} className="text-zinc-400 hover:text-zinc-600">
            <XCircle className="w-6 h-6" />
          </button>
        </div>
        <form onSubmit={handleSubmit} className="p-6 space-y-4">
          <div>
            <label className="block text-xs font-bold text-zinc-500 uppercase mb-1">Username</label>
            <input
              type="text"
              required
              className="w-full px-4 py-2 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-zinc-500 uppercase mb-1">
              {user ? 'New Password (leave blank to keep current)' : 'Password'}
            </label>
            <input
              type="password"
              required={!user}
              className="w-full px-4 py-2 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-zinc-500 uppercase mb-1">Department</label>
            <select
              className="w-full px-4 py-2 rounded-lg border border-zinc-300 outline-none focus:ring-2 focus:ring-indigo-500"
              value={department}
              onChange={(e) => setDepartment(e.target.value)}
            >
              {departments.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
          </div>
          <div>
            <label className="block text-xs font-bold text-zinc-500 uppercase mb-2">Roles</label>
            <div className="grid grid-cols-2 gap-2">
              {availableRoles.map(role => (
                <label key={role.id} className="flex items-center gap-2 p-2 rounded-lg border border-zinc-100 hover:bg-zinc-50 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={selectedRoles.includes(role.name)}
                    onChange={() => toggleRole(role.name)}
                    className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500"
                  />
                  <span className="text-sm text-zinc-700">{role.name}</span>
                </label>
              ))}
            </div>
          </div>
          <div className="pt-4 flex gap-3">
            <button
              type="button"
              onClick={onClose}
              className="flex-1 px-4 py-2 rounded-lg border border-zinc-200 text-zinc-600 font-bold hover:bg-zinc-50 transition-colors"
            >
              Cancel
            </button>
            <button
              type="submit"
              disabled={loading}
              className="flex-1 px-4 py-2 rounded-lg bg-indigo-600 text-white font-bold hover:bg-indigo-700 transition-colors disabled:opacity-50"
            >
              {loading ? 'Saving...' : user ? 'Update User' : 'Create User'}
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
};

const UserManagement = ({ roles: availableRoles, onRolesRefresh }: { roles: Role[], onRolesRefresh: () => void }) => {
  const [users, setUsers] = useState<User[]>([]);
  const [loading, setLoading] = useState(true);
  const [selectedUser, setSelectedUser] = useState<User | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);

  const fetchUsers = async () => {
    try {
      const data = await api.request('/api/users');
      setUsers(data);
    } catch (err) {
      toast.error('Failed to load users');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchUsers();
  }, []);

  if (loading) return <div className="text-center py-12">Loading users...</div>;

  return (
    <div className="space-y-6">
      <RoleManager roles={availableRoles} onRefresh={onRolesRefresh} />
      
      <div className="bg-white rounded-xl border border-zinc-200 overflow-hidden shadow-sm">
        <div className="p-6 border-b border-zinc-100 flex justify-between items-center">
          <h2 className="text-xl font-bold text-zinc-900 flex items-center gap-2">
            <Users className="w-5 h-5 text-indigo-600" />
            User Management
          </h2>
          <button
            onClick={() => {
              setSelectedUser(null);
              setIsModalOpen(true);
            }}
            className="bg-indigo-600 text-white px-4 py-2 rounded-lg font-bold text-sm hover:bg-indigo-700 transition-colors flex items-center gap-2"
          >
            <UserPlus className="w-4 h-4" />
            Add User
          </button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-zinc-50 text-zinc-500 text-xs uppercase font-bold">
              <tr>
                <th className="px-6 py-3">Username</th>
                <th className="px-6 py-3">Department</th>
                <th className="px-6 py-3">Roles</th>
                <th className="px-6 py-3 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-100">
              {users.map((u) => (
                <tr key={u.id} className="hover:bg-zinc-50 transition-colors text-sm">
                  <td className="px-6 py-4 font-medium text-zinc-900">{u.username}</td>
                  <td className="px-6 py-4">
                    <span className="px-2 py-1 bg-zinc-100 text-zinc-600 rounded text-xs font-medium uppercase tracking-wider">
                      {u.department}
                    </span>
                  </td>
                  <td className="px-6 py-4">
                    <div className="flex flex-wrap gap-1">
                      {(u.roles || []).map((r, i) => (
                        <span key={`${r}-${i}`} className={cn(
                          "px-2 py-0.5 rounded-full text-[8px] font-bold uppercase",
                          r === 'admin' ? "bg-indigo-100 text-indigo-700" : "bg-zinc-100 text-zinc-600"
                        )}>
                          {r}
                        </span>
                      ))}
                    </div>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <div className="flex justify-end gap-2">
                      <button 
                        onClick={() => {
                          setSelectedUser(u);
                          setIsModalOpen(true);
                        }}
                        className="p-1.5 text-zinc-400 hover:text-indigo-600 transition-colors"
                        title="Edit User"
                      >
                        <Edit2 className="w-4 h-4" />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {isModalOpen && (
          <UserModal 
            user={selectedUser}
            availableRoles={availableRoles}
            onClose={() => setIsModalOpen(false)}
            onSuccess={() => {
              setIsModalOpen(false);
              fetchUsers();
            }}
          />
        )}
      </AnimatePresence>
    </div>
  );
};

type ProcurementCenterExportTab = 'pr' | 'sr' | 'po' | 'invoice';

function procurementCenterExportTabLabel(tab: ProcurementCenterExportTab): string {
  if (tab === 'pr') return 'PR';
  if (tab === 'sr') return 'SR';
  if (tab === 'po') return 'PO';
  return 'Invoice';
}

function csvEscape(value: unknown): string {
  const s = value == null ? '' : String(value);
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function rowsToCsv(rows: Array<Record<string, unknown>>): string {
  if (!rows.length) return '';
  const headers = Array.from(
    rows.reduce((set, row) => {
      Object.keys(row).forEach((k) => set.add(k));
      return set;
    }, new Set<string>())
  );
  const headerLine = headers.map(csvEscape).join(',');
  const body = rows
    .map((row) => headers.map((h) => csvEscape(row[h])).join(','))
    .join('\r\n');
  return `${headerLine}\r\n${body}`;
}

function downloadCsvFile(fileName: string, csvContent: string): void {
  const blob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/** Export as 3 CSV files: export info, requests summary, and line items. */
async function downloadProcurementCenterExcel(
  rows: WorkflowRequest[],
  tab: ProcurementCenterExportTab,
  filters: { status: string; department: string; search: string; entity: string }
): Promise<void> {
  const tabLabel = procurementCenterExportTabLabel(tab);
  const stamp = new Date().toISOString().slice(0, 10);

  const infoRows = [
    { Key: 'Title', Value: 'Procurement Center export' },
    { Key: 'Exported', Value: formatDateTimeMYT(new Date()) },
    { Key: 'List', Value: tabLabel },
    { Key: 'Search', Value: filters.search.trim() || '—' },
    { Key: 'Status', Value: filters.status.trim() || 'All' },
    { Key: 'Department', Value: filters.department.trim() || '—' },
    { Key: 'Entity filter', Value: filters.entity.trim() || 'All' },
    { Key: 'Row count', Value: rows.length },
  ];

  const summaryRows = rows.map((r) => {
    const m = procurementMoneyTotals(r);
    const curp = procurementCurrencyPrefix(r.currency);
    const lines = (r.line_items || []).map((item: Record<string, unknown>) => {
      const qty =
        item?.Quantity ?? item?.quantity ?? item?.['Max quantity'] ?? item?.['Min quantity'] ?? '';
      const name = String(item?.Item ?? item?.item ?? '').trim();
      return `${qty} × ${name}`.trim();
    }).filter((s) => s.length > 0 && !/^×$/.test(s));
    const linkedPo = r.converted_po_request_id;
    return {
      'Request ID': r.id,
      'Document ID': displayRequestSerial(r),
      'Formatted ID': (r.formatted_id ?? '').toString(),
      Template: r.template_name ?? '',
      Title: r.title ?? '',
      Entity: String(r.entity || '').trim(),
      'Entity name': entityLegalDisplayName(r.entity),
      Department: r.department ?? '',
      Requester: r.requester_name ?? '',
      'Chosen approver ID': r.assigned_approver_id ?? '',
      'Chosen approver': chosenApproverNameLabel(r),
      'Chosen approver (at submit)': String(r.assigned_approver_name_saved ?? '').trim(),
      'Chosen approver designation (at submit)': String(r.assigned_approver_designation_saved ?? '').trim(),
      Status: formatWorkflowRequestStatusLabel(r),
      'Created': r.created_at ?? '',
      Currency: r.currency ?? '',
      Subtotal: m.subtotal,
      'Discount %': Number((m.discountRate * 100).toFixed(4)),
      'Tax %': Number((m.taxRate * 100).toFixed(4)),
      Total: m.total,
      'Total (display)': m.total > 0 ? `${curp}${formatProcurementMoney(m.total)}`.trim() : '',
      'Item count': r.line_items?.length ?? 0,
      'Line summary': lines.join(' | '),
      Supplier: procurementSupplierColumnDisplay(r),
      'Cost center': r.cost_center ?? '',
      Section: r.section ?? '',
      'Linked PO request ID': linkedPo != null && Number(linkedPo) > 0 ? Number(linkedPo) : '',
      Details: (r.details ?? '').toString().slice(0, 8000),
    };
  });

  const lineRows: Record<string, string | number>[] = [];
  for (const r of rows) {
    const docId = displayRequestSerial(r);
    const items = r.line_items || [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i] as Record<string, unknown>;
      const flat: Record<string, string | number> = {
        'Document ID': docId,
        'Request ID': r.id,
        'Request title': (r.title ?? '').toString(),
        'Line #': i + 1,
      };
      if (item && typeof item === 'object') {
        for (const [k, v] of Object.entries(item)) {
          if (k.startsWith('_')) continue;
          let cell: string;
          if (v == null) cell = '';
          else if (typeof v === 'object') cell = JSON.stringify(v);
          else cell = String(v);
          const kl = k.toLowerCase();
          if ((kl.includes('signature') || kl.includes('image')) && cell.length > 120) {
            cell = '[omitted — binary or long text]';
          }
          if (cell.length > 8000) cell = `${cell.slice(0, 8000)}…`;
          flat[k] = cell;
        }
      }
      lineRows.push(flat);
    }
  }
  const lineRowsOut =
    lineRows.length > 0 ? lineRows : [{ Note: 'No line items in this export' }];

  const base = `Procurement_${tabLabel}_${stamp}`.replace(/[/\\?%*:|"<>]/g, '-');
  downloadCsvFile(`${base}_ExportInfo.csv`, rowsToCsv(infoRows));
  downloadCsvFile(`${base}_Requests.csv`, rowsToCsv(summaryRows as Array<Record<string, unknown>>));
  downloadCsvFile(`${base}_LineItems.csv`, rowsToCsv(lineRowsOut as Array<Record<string, unknown>>));
}

const ProcurementCenter = ({ user }: { user: User }) => {
  const [requests, setRequests] = useState<WorkflowRequest[]>([]);
  const [loading, setLoading] = useState(true);
  const [activeSubTab, setActiveSubTab] = useState<'pr' | 'sr' | 'po' | 'invoice'>('pr');
  const [filters, setFilters] = useState({
    status: '',
    department: '',
    search: '',
    entity: ''
  });
  const filtersRef = useRef(filters);
  filtersRef.current = filters;
  const [debouncedProcurementSearch, setDebouncedProcurementSearch] = useState(() => filters.search);
  useEffect(() => {
    const t = window.setTimeout(() => setDebouncedProcurementSearch(filters.search), 400);
    return () => window.clearTimeout(t);
  }, [filters.search]);
  const [tableSortBy, setTableSortBy] = useState<
    'id' | 'type' | 'entity' | 'requester' | 'chosenApprover' | 'items' | 'supplier' | 'total' | 'status' | 'date'
  >('date');
  const [tableSortDirection, setTableSortDirection] = useState<'asc' | 'desc'>('desc');
  const [selectedRequest, setSelectedRequest] = useState<WorkflowRequest | null>(null);
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [approvals, setApprovals] = useState<RequestApproval[]>([]);
  const [isEditing, setIsEditing] = useState(false);
  const [editData, setEditData] = useState({
    title: '',
    details: '',
    line_items: [] as any[],
    tax_rate: 18,
    discount_rate: 0,
    currency: '',
    cost_center: '',
    section: '',
    suggested_supplier: '',
    requester_username: '',
    requester_name: '',
  });
  const [editAttachmentKeep, setEditAttachmentKeep] = useState<Attachment[]>([]);
  const [editAttachmentAdd, setEditAttachmentAdd] = useState<{ id: string; name: string; type: string; data: string }[]>([]);
  const [viewingPdf, setViewingPdf] = useState<{ url: string; fileName: string } | null>(null);
  const [detailsViewMode, setDetailsViewMode] = useState<'details' | 'pdf'>('details');
  const [detailsPdfPreview, setDetailsPdfPreview] = useState<{ url: string; fileName: string } | null>(null);
  const [convertPoModal, setConvertPoModal] = useState<ConvertPoModalTarget | null>(null);
  const [purchasingDecisionModal, setPurchasingDecisionModal] = useState<PurchasingDecisionModalTarget | null>(null);
  const {
    options: costCenterOptions,
    loading: costCentersLoading,
    failed: costCentersLoadFailed,
    fetchError: costCentersFetchError,
  } = useCostCenterOptions();

  const isPurchasing = user.roles?.some(r => r.toLowerCase() === 'purchasing');
  const isAdmin = user.roles?.some(r => r.toLowerCase() === 'admin') || user.permissions?.includes('admin');
  const canEdit = selectedRequest && (
    (isWorkflowRequestPending(selectedRequest) &&
      (user.id === selectedRequest.requester_id ||
        user.permissions?.includes('admin') ||
        (isPurchasing && isPO_Only(selectedRequest)))) ||
    (isWorkflowRequestRejected(selectedRequest) &&
      isPR_Only(selectedRequest) &&
      user.id === selectedRequest.requester_id) ||
    (isWorkflowRequestFullyApproved(selectedRequest) && isPR_Only(selectedRequest) && isPurchasing)
  );
  const canRequesterCancelSelected =
    !!selectedRequest && canRequesterCancelPendingRequest(selectedRequest, user, approvals);

  const fetchProcurementRequests = useCallback(
    async (immediateSearch = false): Promise<WorkflowRequest[]> => {
      const f = filtersRef.current;
      const searchForQuery = immediateSearch ? f.search : debouncedProcurementSearch;
      setLoading(true);
      try {
        // Do not inherit the header entity: omit `entity` from the query to load all rows for
        // `user.entities` (server uses IN(...)). The entity filter dropdown uses `user.entities`, not
        // only entities present in the current page of results.
        const mergedFilters = {
          ...f,
          search: searchForQuery,
          entity: String(f.entity || '').trim(),
        };
        const params = new URLSearchParams();
        for (const [k, v] of Object.entries(mergedFilters)) {
          const s = v == null ? '' : String(v).trim();
          if (s !== '') params.set(k, s);
        }
        const qs = params.toString();
        const data = await api.request(`/api/procurement/requests${qs ? `?${qs}` : ''}`, { skipEntity: true });
        setRequests(data);
        return data;
      } catch (err) {
        toast.error('Failed to fetch procurement requests');
        return [];
      } finally {
        setLoading(false);
      }
    },
    [debouncedProcurementSearch]
  );

  const fetchDetails = async (id: number) => {
    try {
      const [attData, appData] = await Promise.all([
        api.request(`/api/workflow-requests/${id}/attachments`),
        api.request(`/api/workflow-requests/${id}/approvals`)
      ]);
      setAttachments(attData);
      setApprovals(appData);
    } catch (err) {
      console.error(err);
    }
  };

  /** List payload omits heavy fields; merge full row for modal (signature, details, steps). */
  const openProcurementRequestDetail = async (lite: WorkflowRequest) => {
    const ent = String(lite.entity || '').trim();
    if (ent) api.setActiveEntity(ent);
    else api.setActiveEntity(null);
    setSelectedRequest(lite);
    setDetailsViewMode('details');
    setDetailsPdfPreview((prev) => {
      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
      return null;
    });
    void fetchDetails(lite.id);
    try {
      const full = (await api.request(`/api/workflow-requests/${lite.id}`)) as WorkflowRequest;
      setSelectedRequest((prev) => (prev && prev.id === lite.id ? { ...prev, ...full } : prev));
    } catch (err: any) {
      toast.error(err?.message || 'Could not load full request details');
    }
  };

  const handleEditAttachmentAdd = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;
    Array.from(files).forEach((file: File) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        setEditAttachmentAdd((prev) => [
          ...prev,
          {
            id: Math.random().toString(36).substr(2, 9),
            name: file.name,
            type: file.type,
            data: event.target?.result as string,
          },
        ]);
      };
      reader.readAsDataURL(file);
    });
    e.target.value = '';
  };

  const handleSaveEdit = async () => {
    if (!selectedRequest) return;
    if (
      (isPRRequest(selectedRequest) || isPO_Only(selectedRequest) || isSRRequest(selectedRequest)) &&
      !String(editData.currency).trim()
    ) {
      return toast.error('Please select a currency.');
    }
    if (
      isSRRequest(selectedRequest) &&
      !SECTION_OPTIONS.includes(String(editData.section || "").trim().toUpperCase() as (typeof SECTION_OPTIONS)[number])
    ) {
      return toast.error('Please select a section.');
    }
    if (isPRRequest(selectedRequest) && !String(editData.suggested_supplier ?? '').trim()) {
      return toast.error('Please enter the suggested supplier for this PR.');
    }
    if (selectedRequest.attachments_required && (editAttachmentKeep.length + editAttachmentAdd.length) === 0) {
      return toast.error('This workflow requires at least one attachment.');
    }
    setLoading(true);
    try {
      const detailsOut = isProcurementPRorPORequest(selectedRequest) ? '' : editData.details;
      const taxUnit = procurementPercentToUnitRate(Number(editData.tax_rate) || 0);
      const discUnit = procurementPercentToUnitRate(Number(editData.discount_rate) || 0);
      const sectionOut = isSRRequest(selectedRequest) ? sectionPayloadFromSelection(editData.section) : editData.section;
      const normalizedLineItems = normalizeProcurementCatalogLineItems(
        editData.line_items,
        costCenterOptions,
        isSRRequest(selectedRequest)
      );
      const curOut =
        isPRRequest(selectedRequest) || isPO_Only(selectedRequest) || isSRRequest(selectedRequest)
          ? String(editData.currency).trim()
          : undefined;
      await api.request(`/api/workflow-requests/${selectedRequest.id}`, {
        method: 'PATCH',
        body: JSON.stringify({
          ...editData,
          line_items: normalizedLineItems,
          tax_rate: taxUnit,
          discount_rate: discUnit,
          details: detailsOut,
          currency: curOut,
          section: sectionOut,
          attachment_keep_ids: editAttachmentKeep.map((a) => a.id),
          attachments_add: editAttachmentAdd.map((a) => ({ name: a.name, type: a.type, data: a.data })),
          ...(isPRRequest(selectedRequest)
            ? { suggested_supplier: String(editData.suggested_supplier ?? '').trim() }
            : {}),
        }),
      });
      toast.success('Request updated');
      setIsEditing(false);
      setEditAttachmentAdd([]);
      const refreshed = await fetchProcurementRequests(true);
      const nextSel = refreshed.find((x) => x.id === selectedRequest.id);
      if (nextSel) {
        setSelectedRequest(nextSel);
      } else {
        setSelectedRequest({
          ...selectedRequest,
          ...editData,
          line_items: normalizedLineItems,
          tax_rate: taxUnit,
          discount_rate: discUnit,
          currency: curOut !== undefined ? curOut : selectedRequest.currency,
          details: detailsOut,
          section: sectionOut,
          ...(isPRRequest(selectedRequest)
            ? { suggested_supplier: String(editData.suggested_supplier ?? '').trim() }
            : {}),
        });
      }
      await fetchDetails(selectedRequest.id);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleResubmit = async (id: number) => {
    setLoading(true);
    try {
      await api.request(`/api/workflow-requests/${id}/resubmit`, {
        method: 'POST',
      });
      toast.success('Request resubmitted successfully!');
      void fetchProcurementRequests(true);
      setSelectedRequest(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleRequesterCancel = async (id: number) => {
    const promptResult = window.prompt('Optional cancellation reason (for audit):', '');
    if (promptResult === null) return;
    const reason = promptResult.trim();
    const confirmed = window.confirm('Cancel this request? This action is final and cannot be resubmitted.');
    if (!confirmed) return;
    setLoading(true);
    try {
      await api.request(`/api/workflow-requests/${id}/cancel`, {
        method: 'POST',
        body: JSON.stringify({ comment: reason }),
      });
      toast.success('Request cancelled');
      void fetchProcurementRequests(true);
      setSelectedRequest(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handlePrintPR = (request: WorkflowRequest) => {
    const merged = mergeRequestApprovalsForPdf(request, selectedRequest?.id, approvals);
    const preview = createWorkflowRequestPdfPreview(merged);
    void persistGeneratedProcurementFormPdf(request.id, preview.pdfDataUrl);
    setViewingPdf((prev) => {
      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
      return { url: preview.url, fileName: preview.fileName };
    });
    toast.success('PDF ready to view');
  };

  const handleShowRequestPdfInline = (request: WorkflowRequest) => {
    const merged = mergeRequestApprovalsForPdf(request, selectedRequest?.id, approvals);
    const preview = createWorkflowRequestPdfPreview(merged);
    void persistGeneratedProcurementFormPdf(request.id, preview.pdfDataUrl);
    setDetailsPdfPreview((prev) => {
      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
      return { url: preview.url, fileName: preview.fileName };
    });
    setDetailsViewMode('pdf');
  };

  const handleGeneratePODraft = (request: WorkflowRequest) => {
    const doc = new jsPDF();
    const margin = 20;
    let y = 20;

    // Header
    doc.setFontSize(20);
    doc.setTextColor(79, 70, 229); // Indigo-600
    doc.text('PURCHASE ORDER DRAFT', margin, y);
    y += 10;

    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Generated on: ${formatDateTimeMYT(new Date())}`, margin, y);
    y += 15;

    // Request Info
    doc.setFontSize(12);
    doc.setTextColor(0);
    doc.setFont('helvetica', 'bold');
    doc.text('Request Information', margin, y);
    y += 7;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.text(`PR ID: ${displayRequestSerial(request)}`, margin, y);
    y += 5;
    doc.text(`Title: ${request.title}`, margin, y);
    y += 5;
    doc.text(`Department: ${request.department}`, margin, y);
    y += 5;
    doc.text(`Entity: ${entityLegalDisplayName(request.entity)} (${request.entity?.trim() || '-'})`, margin, y);
    y += 15;

    // Line Items Table
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text('Order Items', margin, y);
    y += 7;

    const headers = ['Item', 'Qty', 'Unit Price', 'Total'];
    const colWidths = [80, 20, 35, 35];
    
    // Table Header
    doc.setFillColor(244, 244, 245);
    doc.rect(margin, y - 5, 170, 7, 'F');
    doc.setFontSize(9);
    let x = margin;
    headers.forEach((h, i) => {
      doc.text(h, x + 2, y);
      x += colWidths[i];
    });
    y += 7;

    // Table Rows
    doc.setFont('helvetica', 'normal');
    request.line_items?.forEach((item) => {
      const qty = parseFloat(item['Quantity'] || '0');
      const price = parseFloat(item['Unit Price'] || '0');
      const rowTot = qty * price;

      x = margin;
      doc.text(String(item['Item'] || '-'), x + 2, y, { maxWidth: colWidths[0] - 4 });
      doc.text(String(qty), x + colWidths[0] + 2, y);
      doc.text(`${procurementCurrencyPrefix(request.currency)}${formatProcurementMoney(price)}`.trim(), x + colWidths[0] + colWidths[1] + 2, y);
      doc.text(`${procurementCurrencyPrefix(request.currency)}${formatProcurementMoney(rowTot)}`.trim(), x + colWidths[0] + colWidths[1] + colWidths[2] + 2, y);
      y += 7;
    });

    y += 5;
    const m = procurementMoneyTotals(request);
    const curp = procurementCurrencyPrefix(request.currency);
    doc.setFont('helvetica', 'bold');
    doc.text(`Subtotal: ${curp}${formatProcurementMoney(m.subtotal)}`.trim(), 110, y);
    y += 5;
    doc.text(`Discount (${(m.discountRate * 100).toFixed(0)}%): ${curp}${formatProcurementMoney(m.discountAmount)}`.trim(), 110, y);
    y += 5;
    doc.text(`${procurementTaxLabelForEntity(request.entity)} (${(m.taxRate * 100).toFixed(0)}%): ${curp}${formatProcurementMoney(m.taxAmount)}`.trim(), 110, y);
    y += 7;
    doc.setFontSize(12);
    doc.text(`TOTAL AMOUNT: ${curp}${formatProcurementMoney(m.total)}`.trim(), 110, y);

    doc.save(`PO_Draft_${displayRequestSerial(request)}.pdf`);
    toast.success('PO Draft PDF generated');
  };

  const handleUploadRealPoAttachment = async (file: File) => {
    if (!selectedRequest || !isPO_Only(selectedRequest)) return;
    setLoading(true);
    try {
      const data = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          if (typeof reader.result === 'string') resolve(reader.result);
          else reject(new Error('Failed to read selected file'));
        };
        reader.onerror = () => reject(new Error('Failed to read selected file'));
        reader.readAsDataURL(file);
      });
      await api.request(`/api/workflow-requests/${selectedRequest.id}`, {
        method: 'PATCH',
        body: JSON.stringify({
          attachments_add: [{ name: file.name, type: file.type || 'application/octet-stream', data }],
        }),
      });
      toast.success('PO document uploaded');
      await fetchDetails(selectedRequest.id);
      await fetchProcurementRequests(true);
    } catch (err: any) {
      toast.error(err?.message || 'Failed to upload PO document');
    } finally {
      setLoading(false);
    }
  };

  const openConvertPoModal = (r: WorkflowRequest) => {
    setConvertPoModal({
      id: r.id,
      title: r.title || '',
      prSerial: r.formatted_id ? toUpperSerial(r.formatted_id) : `#${r.id}`,
      entityCode: String(r.entity || '').trim() || '—',
    });
  };

  const handleConvertToPOConfirm = async (poNumber: string, poUpload?: ConvertPoUploadPayload) => {
    if (!convertPoModal) return;
    setLoading(true);
    try {
      const result = await api.request(`/api/workflow-requests/${convertPoModal.id}/convert-to-po`, {
        method: 'POST',
        body: JSON.stringify({ po_number: poNumber, po_upload: poUpload || null }),
      });
      toast.success(
        result?.merged_into_existing
          ? `PR appended to PO: ${toUpperSerial(result.formatted_id)}`
          : `PO created: ${toUpperSerial(result.formatted_id)}`
      );
      setConvertPoModal(null);
      void fetchProcurementRequests(true);
      setSelectedRequest(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  const openPurchasingDecisionModal = (r: WorkflowRequest, decision: 'rejected' | 'cancelled') => {
    const ent = String(r.entity || '').trim();
    if (ent) api.setActiveEntity(ent);
    setPurchasingDecisionModal({
      id: r.id,
      decision,
      prSerial: r.formatted_id ? toUpperSerial(r.formatted_id) : `#${r.id}`,
      title: r.title || '',
      entity: ent,
    });
  };

  const handlePurchasingFinalDecision = async (id: number, decision: 'rejected' | 'cancelled', comment: string, entity: string) => {
    setLoading(true);
    try {
      const ent = String(entity || '').trim();
      if (ent) api.setActiveEntity(ent);
      await api.request(`/api/workflow-requests/${id}/purchasing-decision`, {
        method: 'POST',
        body: JSON.stringify({
          decision,
          comment,
        }),
      });
      toast.success(decision === 'cancelled' ? 'PR cancelled' : 'PR rejected by purchasing');
      void fetchProcurementRequests(true);
      setSelectedRequest(null);
      setPurchasingDecisionModal(null);
    } catch (err: any) {
      toast.error(err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void fetchProcurementRequests(false);
  }, [fetchProcurementRequests, filters.status, filters.department, filters.entity, debouncedProcurementSearch]);

  const procurementEntityFilterOptions = Array.from(
    new Set((user.entities || []).map((e) => String(e).trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

  const filteredRequests = requests.filter((r) => {
    const templateName = r.template_name.toLowerCase();
    if (activeSubTab === 'pr') {
      return (
        !isPO_Only(r) &&
        !isSRRequest(r) &&
        (templateName.includes('purchase request') || templateName.includes('pr'))
      );
    }
    if (activeSubTab === 'sr') return isSRRequest(r);
    if (activeSubTab === 'po') return isPO_Only(r);
    if (activeSubTab === 'invoice') return templateName.includes('invoice');
    return true;
  });

  const toggleProcurementSort = (
    next:
      | 'id'
      | 'type'
      | 'entity'
      | 'requester'
      | 'chosenApprover'
      | 'items'
      | 'supplier'
      | 'total'
      | 'status'
      | 'date'
  ) => {
    if (tableSortBy === next) {
      setTableSortDirection((d) => (d === 'asc' ? 'desc' : 'asc'));
      return;
    }
    setTableSortBy(next);
    setTableSortDirection(next === 'date' ? 'desc' : 'asc');
  };

  const sortedFilteredRequests = [...filteredRequests].sort((a, b) => {
    const dir = tableSortDirection === 'asc' ? 1 : -1;
    const byText = (x: string, y: string) => x.localeCompare(y, undefined, { sensitivity: 'base' }) * dir;
    const supplierOf = (r: WorkflowRequest) => prSuggestedSupplierDisplay(r);
    if (tableSortBy === 'date') return (new Date(a.created_at).getTime() - new Date(b.created_at).getTime()) * dir;
    if (tableSortBy === 'id') return byText(displayRequestSerial(a), displayRequestSerial(b));
    if (tableSortBy === 'type') return byText(`${a.template_name} ${a.title || ''}`, `${b.template_name} ${b.title || ''}`);
    if (tableSortBy === 'entity') return byText(String(a.entity || ''), String(b.entity || ''));
    if (tableSortBy === 'requester') return byText(String(a.requester_name || ''), String(b.requester_name || ''));
    if (tableSortBy === 'chosenApprover') return byText(chosenApproverNameLabel(a), chosenApproverNameLabel(b));
    if (tableSortBy === 'items') return ((a.line_items?.length || 0) - (b.line_items?.length || 0)) * dir;
    if (tableSortBy === 'supplier') return byText(supplierOf(a), supplierOf(b));
    if (tableSortBy === 'total') return (procurementMoneyTotals(a).total - procurementMoneyTotals(b).total) * dir;
    return byText(formatWorkflowRequestStatusLabel(a), formatWorkflowRequestStatusLabel(b));
  });

  const procurementSortIndicator = (
    key:
      | 'id'
      | 'type'
      | 'entity'
      | 'requester'
      | 'chosenApprover'
      | 'items'
      | 'supplier'
      | 'total'
      | 'status'
      | 'date'
  ) => (tableSortBy === key ? (tableSortDirection === 'asc' ? ' ↑' : ' ↓') : '');

  const emptyProcurementListMessage =
    activeSubTab === 'pr'
      ? 'No purchase requests found.'
      : activeSubTab === 'sr'
        ? 'No stock requisitions found.'
        : activeSubTab === 'po'
          ? 'No purchase orders found.'
          : 'No invoice requests found.';

  const handleExportProcurementExcel = async () => {
    if (sortedFilteredRequests.length === 0) {
      toast.error('No rows to export for the current tab and filters.');
      return;
    }
    try {
      await downloadProcurementCenterExcel(sortedFilteredRequests, activeSubTab, filters);
      toast.success('Excel file downloaded');
    } catch (err: any) {
      toast.error(err?.message || 'Export failed');
    }
  };

  return (
    <div className="h-full min-h-0 flex flex-col gap-6">
      {/* Header & Tabs */}
      <div className="bg-white p-6 rounded-2xl border border-zinc-200 shadow-sm">
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 bg-indigo-100 rounded-xl flex items-center justify-center">
              <ShoppingCart className="w-6 h-6 text-indigo-600" />
            </div>
            <div>
              <h1 className="text-2xl font-bold text-zinc-900">Procurement Center</h1>
              <p className="text-sm text-zinc-500">Manage and monitor all procurement-related workflows</p>
            </div>
          </div>
        </div>

        <div className="flex flex-wrap gap-2 p-1 bg-zinc-100 rounded-xl w-full max-w-full">
          <button
            onClick={() => setActiveSubTab('pr')}
            className={cn(
              "px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-bold transition-all flex items-center gap-2",
              activeSubTab === 'pr' ? "bg-white text-indigo-600 shadow-sm" : "text-zinc-500 hover:text-zinc-700"
            )}
          >
            <ClipboardList className="w-4 h-4 shrink-0" />
            <span className="whitespace-nowrap">Purchase Requests (PR)</span>
          </button>
          <button
            onClick={() => setActiveSubTab('sr')}
            className={cn(
              "px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-bold transition-all flex items-center gap-2",
              activeSubTab === 'sr' ? "bg-white text-indigo-600 shadow-sm" : "text-zinc-500 hover:text-zinc-700"
            )}
          >
            <Warehouse className="w-4 h-4 shrink-0" />
            <span className="whitespace-nowrap">Stock Requisition (SR)</span>
          </button>
          <button
            onClick={() => setActiveSubTab('po')}
            className={cn(
              "px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-bold transition-all flex items-center gap-2",
              activeSubTab === 'po' ? "bg-white text-indigo-600 shadow-sm" : "text-zinc-500 hover:text-zinc-700"
            )}
          >
            <Package className="w-4 h-4 shrink-0" />
            <span className="whitespace-nowrap">Purchase Orders (PO)</span>
          </button>
          <button
            onClick={() => setActiveSubTab('invoice')}
            className={cn(
              "px-3 sm:px-4 py-2 rounded-lg text-xs sm:text-sm font-bold transition-all flex items-center gap-2",
              activeSubTab === 'invoice' ? "bg-white text-indigo-600 shadow-sm" : "text-zinc-500 hover:text-zinc-700"
            )}
          >
            <Receipt className="w-4 h-4 shrink-0" />
            <span className="whitespace-nowrap">Invoice Requests</span>
          </button>
        </div>
      </div>

      {/* Filters */}
      <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm flex flex-wrap gap-4 items-center">
        <div className="flex-1 min-w-[200px] relative">
          <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" />
          <input
            type="text"
            placeholder="Search by ID or Title..."
            className="w-full pl-10 pr-4 py-2 rounded-xl border border-zinc-200 outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
            value={filters.search}
            onChange={(e) => setFilters({ ...filters, search: e.target.value })}
          />
        </div>
        <select
          className="px-4 py-2 rounded-xl border border-zinc-200 outline-none focus:ring-2 focus:ring-indigo-500 text-sm bg-white"
          value={filters.status}
          onChange={(e) => setFilters({ ...filters, status: e.target.value })}
        >
          <option value="">All Statuses</option>
          <option value="pending">Pending</option>
          <option value="approved">Approved</option>
          <option value="rejected">Rejected</option>
          <option value="cancelled">Cancelled</option>
        </select>
        <input
          type="text"
          placeholder="Department..."
          className="px-4 py-2 rounded-xl border border-zinc-200 outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
          value={filters.department}
          onChange={(e) => setFilters({ ...filters, department: e.target.value })}
        />
        <select
          className="px-4 py-2 rounded-xl border border-zinc-200 outline-none focus:ring-2 focus:ring-indigo-500 text-sm bg-white"
          value={filters.entity}
          onChange={(e) => setFilters({ ...filters, entity: e.target.value })}
        >
          <option value="">All Entities</option>
          {procurementEntityFilterOptions.map((ent) => (
            <option key={ent} value={ent}>{ent}</option>
          ))}
        </select>
        <button
          type="button"
          onClick={handleExportProcurementExcel}
          disabled={loading || sortedFilteredRequests.length === 0}
          title={
            sortedFilteredRequests.length === 0
              ? 'Nothing to export — adjust filters or refresh'
              : 'Download current list as .xlsx (summary + line items)'
          }
          className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-xl text-sm font-bold hover:bg-emerald-700 transition-colors disabled:opacity-40 disabled:pointer-events-none"
        >
          <FileSpreadsheet className="w-4 h-4" />
          Export to Excel
        </button>
        <button 
          onClick={() => void fetchProcurementRequests(true)}
          className="px-4 py-2 bg-zinc-900 text-white rounded-xl text-sm font-bold hover:bg-zinc-800 transition-colors"
        >
          Refresh
        </button>
      </div>

      {/* Table */}
      <div className="bg-white rounded-2xl border border-zinc-200 shadow-sm overflow-hidden flex-1 min-h-0">
        <div className="w-full h-full overflow-auto min-w-0">
        <table className="w-full min-w-[1200px] table-auto text-left border-collapse text-[11px] sm:text-xs">
          <thead>
            <tr className="bg-zinc-50 border-b border-zinc-200 sticky top-0 z-[1] shadow-[0_1px_0_0_rgb(228_228_231)]">
              <th className="px-3 py-2 min-w-[10.5rem] whitespace-nowrap text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('id')} className="hover:text-zinc-700 transition-colors text-left">
                  ID{procurementSortIndicator('id')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[13rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('type')} className="hover:text-zinc-700 transition-colors text-left">
                  Type{procurementSortIndicator('type')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[7rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('entity')} className="hover:text-zinc-700 transition-colors text-left">
                  Entity{procurementSortIndicator('entity')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[8.5rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('requester')} className="hover:text-zinc-700 transition-colors text-left">
                  Requester{procurementSortIndicator('requester')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[8.5rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('chosenApprover')} className="hover:text-zinc-700 transition-colors text-left">
                  Chosen approver{procurementSortIndicator('chosenApprover')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[10rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('items')} className="hover:text-zinc-700 transition-colors text-left">
                  Items{procurementSortIndicator('items')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[8.5rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('supplier')} className="hover:text-zinc-700 transition-colors text-left">
                  Supplier{procurementSortIndicator('supplier')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[5.5rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('total')} className="hover:text-zinc-700 transition-colors text-left">
                  Total{procurementSortIndicator('total')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[6.5rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('status')} className="hover:text-zinc-700 transition-colors text-left">
                  Status{procurementSortIndicator('status')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[5.5rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider align-bottom">
                <button type="button" onClick={() => toggleProcurementSort('date')} className="hover:text-zinc-700 transition-colors text-left">
                  Date{procurementSortIndicator('date')}
                </button>
              </th>
              <th className="px-3 py-2 min-w-[9rem] text-[10px] sm:text-xs font-bold text-zinc-500 uppercase tracking-wider text-right align-bottom">Action</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-zinc-100">
            {loading ? (
              <tr>
                <td colSpan={11} className="px-3 py-12 text-center text-zinc-400 italic">Loading requests...</td>
              </tr>
            ) : sortedFilteredRequests.length === 0 ? (
              <tr>
                <td colSpan={11} className="px-3 py-12 text-center text-zinc-400 italic">{emptyProcurementListMessage}</td>
              </tr>
            ) : (
              sortedFilteredRequests.map((r) => {
                const totalAmount = procurementMoneyTotals(r).total;

                return (
                  <tr 
                    key={r.id} 
                    className="hover:bg-zinc-50 transition-colors cursor-pointer group align-top"
                    onClick={() => {
                      void openProcurementRequestDetail(r);
                    }}
                  >
                    <td className="px-3 py-2 align-top font-mono text-zinc-500 tabular-nums whitespace-nowrap min-w-[10.5rem]">
                      #{displayRequestSerial(r)}
                    </td>
                    <td className="px-3 py-2 align-top min-w-[13rem] max-w-md">
                      <p className="font-bold text-zinc-900 leading-snug break-words">{r.template_name}</p>
                      <p className="text-zinc-500 mt-0.5 break-words text-[10px] sm:text-xs leading-snug">{r.title}</p>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[7rem]">
                      <div className="flex flex-col gap-0.5 break-words leading-tight">
                        <span className="text-[10px] sm:text-xs font-bold text-indigo-600 uppercase tracking-wide">
                          {String(r.entity || '-')}
                        </span>
                        <span className="text-[10px] text-zinc-500">
                          {entityLegalDisplayName(r.entity)}
                        </span>
                      </div>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[8.5rem]">
                      <p className="text-zinc-700 font-medium break-words leading-tight">{r.requester_name}</p>
                      <p className="text-[9px] text-zinc-400 font-bold uppercase">{r.department}</p>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[8.5rem]">
                      <p className="text-zinc-700 font-medium break-words leading-tight">
                        {requestHasChosenApprover(r) ? chosenApproverNameLabel(r) : '—'}
                      </p>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[10rem]">
                      <div className="space-y-0.5">
                        {r.line_items?.slice(0, 2).map((item, idx) => (
                          <div key={item.id || idx} className="text-zinc-700 leading-tight break-words">
                            <span className="font-bold text-indigo-600">{item.Quantity || item.quantity || 0}x</span>{' '}
                            {item.Item || item.item || 'Item'}
                          </div>
                        ))}
                        {r.line_items && r.line_items.length > 2 && (
                          <p className="text-[9px] text-zinc-400 font-bold uppercase">+{r.line_items.length - 2} more</p>
                        )}
                        {(!r.line_items || r.line_items.length === 0) && (
                          <p className="text-zinc-400 italic">No items</p>
                        )}
                      </div>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[8.5rem]">
                      <p className="text-zinc-700 font-medium leading-tight break-words">
                        {procurementSupplierColumnDisplay(r)}
                      </p>
                    </td>
                    <td className="px-3 py-2 align-top whitespace-nowrap min-w-[5.5rem]">
                      <p className="font-bold text-zinc-900 tabular-nums">
                        {totalAmount > 0 ? `${procurementCurrencyPrefix(r.currency)}${formatProcurementMoney(totalAmount)}`.trim() : '-'}
                      </p>
                    </td>
                    <td className="px-3 py-2 align-top min-w-[6.5rem]">
                      <span className={cn(
                        'inline-flex text-[9px] sm:text-[10px] uppercase font-bold px-1.5 py-0.5 rounded-full leading-tight',
                        workflowRequestStatusBadgeClass(r)
                      )}>
                        {formatWorkflowRequestStatusLabel(r)}
                      </span>
                    </td>
                    <td className="px-3 py-2 align-top text-zinc-500 whitespace-nowrap min-w-[5.5rem]">
                      {formatDateMYT(r.created_at)}
                    </td>
                    <td className="px-3 py-2 align-top text-right min-w-[9rem]">
                      <div className="flex flex-col items-end gap-1">
                        {canShowConvertPRToPO(r, user) && (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              openConvertPoModal(r);
                            }}
                            title="Convert to PO"
                            className="inline-flex items-center justify-center gap-1 px-2 py-1 bg-indigo-600 text-white rounded-md font-bold text-[10px] hover:bg-indigo-700 transition-colors shadow-sm whitespace-nowrap"
                          >
                            <RefreshCw className="w-3 h-3 shrink-0" />
                            <span className="hidden lg:inline">Convert to PO</span>
                            <span className="lg:hidden">→ PO</span>
                          </button>
                        )}
                        {canShowPurchasingFinalDecision(r, user) && (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              openPurchasingDecisionModal(r, 'rejected');
                            }}
                            className="inline-flex items-center justify-center gap-1 px-2 py-1 bg-rose-50 text-rose-700 rounded-md font-bold text-[10px] hover:bg-rose-100 transition-all border border-rose-200 shadow-sm whitespace-nowrap"
                            title="Reject returns PR to requester for edit + resubmit"
                          >
                            <XCircle className="w-3 h-3 shrink-0" />
                            <span className="hidden xl:inline">Reject PR</span>
                            <span className="xl:hidden">Reject</span>
                          </button>
                        )}
                        {canShowPurchasingFinalDecision(r, user) && (
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              openPurchasingDecisionModal(r, 'cancelled');
                            }}
                            title="Cancellation is final and cannot be resubmitted"
                            className="inline-flex items-center justify-center gap-1 px-2 py-1 bg-rose-50 text-rose-700 rounded-md font-bold text-[10px] hover:bg-rose-100 transition-all border border-rose-200 shadow-sm whitespace-nowrap"
                          >
                            <XCircle className="w-3 h-3 shrink-0" />
                            <span className="hidden xl:inline">Cancel PR</span>
                            <span className="xl:hidden">Cancel</span>
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            void openProcurementRequestDetail(r);
                          }}
                          title="Request details"
                          className="inline-flex items-center justify-center gap-1 px-2 py-1 bg-indigo-50 text-indigo-600 rounded-md font-bold text-[10px] hover:bg-indigo-100 transition-all border border-indigo-100 shadow-sm whitespace-nowrap"
                        >
                          <FileText className="w-3 h-3 shrink-0" />
                          Details
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
        </div>
      </div>

      {/* Details Modal */}
      <AnimatePresence>
        {selectedRequest && (
          <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white w-full max-w-[96vw] rounded-2xl shadow-2xl overflow-hidden flex flex-col h-[94vh]"
            >
              <div className="p-6 border-b border-zinc-100 flex items-center justify-between bg-zinc-50 shrink-0">
                <div>
                  <h2 className="text-xl font-bold text-zinc-900">
                    {selectedRequest.formatted_id ? `[${toUpperSerial(selectedRequest.formatted_id)}] ${selectedRequest.title}` : selectedRequest.title}
                  </h2>
                  <p className="text-sm text-zinc-500">Template: {selectedRequest.template_name} • Dept: {selectedRequest.department} • Entity: <span className="font-bold text-indigo-600">{entityLegalDisplayName(selectedRequest.entity)}</span> <span className="text-zinc-400">({selectedRequest.entity?.trim() || '-'})</span></p>
                </div>
                <div className="flex items-center gap-4">
                  {canEdit && !isEditing && (
                    <button 
                      onClick={() => {
                        setEditData({ 
                          title: selectedRequest.title, 
                          details: selectedRequest.details, 
                          line_items: normalizeLineItemsForDateInputs(
                            [...(selectedRequest.line_items || [])],
                            procurementGridColumns(selectedRequest)
                          ),
                          tax_rate: procurementUnitRateToPercent(selectedRequest.tax_rate, 0.18),
                          discount_rate: procurementUnitRateToPercent(selectedRequest.discount_rate, 0),
                          currency: selectedRequest.currency?.trim() ?? '',
                          cost_center: selectedRequest.cost_center || '',
                          section: sectionSelectionFromStored(selectedRequest.section),
                          suggested_supplier: prSuggestedSupplierDisplay(selectedRequest),
                          requester_username: String(selectedRequest.requester_username ?? '').trim(),
                          requester_name: String(selectedRequest.requester_name ?? '').trim(),
                        });
                        setEditAttachmentKeep(attachments);
                        setEditAttachmentAdd([]);
                        setIsEditing(true);
                      }}
                      className="flex items-center gap-2 px-3 py-1.5 rounded-lg border border-indigo-200 text-indigo-600 text-xs font-bold hover:bg-indigo-50 transition-all"
                    >
                      <Edit2 className="w-3.5 h-3.5" />
                      Edit Request
                    </button>
                  )}
                  {canShowPurchasingApprovedPRHeaderActions(selectedRequest, user) && !isEditing && (
                    <>
                      {canShowConvertPRToPO(selectedRequest, user) && (
                        <button 
                          onClick={() => {
                            setEditData({ 
                              title: selectedRequest.title, 
                              details: selectedRequest.details, 
                              line_items: normalizeLineItemsForDateInputs(
                                [...(selectedRequest.line_items || [])],
                                procurementGridColumns(selectedRequest)
                              ),
                              tax_rate: procurementUnitRateToPercent(selectedRequest.tax_rate, 0.18),
                              discount_rate: procurementUnitRateToPercent(selectedRequest.discount_rate, 0),
                              currency: selectedRequest.currency?.trim() ?? '',
                              cost_center: selectedRequest.cost_center || '',
                              section: sectionSelectionFromStored(selectedRequest.section),
                              suggested_supplier: prSuggestedSupplierDisplay(selectedRequest),
                              requester_username: String(selectedRequest.requester_username ?? '').trim(),
                              requester_name: String(selectedRequest.requester_name ?? '').trim(),
                            });
                            setEditAttachmentKeep(attachments);
                            setEditAttachmentAdd([]);
                            setIsEditing(true);
                          }}
                          className="flex items-center gap-2 px-4 py-2 bg-emerald-50 border border-emerald-200 rounded-xl text-sm font-bold text-emerald-700 hover:bg-emerald-100 transition-colors"
                        >
                          <Edit2 className="w-4 h-4" />
                          Edit & Add PO Info
                        </button>
                      )}
                      <button 
                        onClick={() => handleGeneratePODraft(selectedRequest)}
                        className="flex items-center gap-2 px-4 py-2 bg-amber-50 border border-amber-200 rounded-xl text-sm font-bold text-amber-700 hover:bg-amber-100 transition-colors"
                      >
                        <Download className="w-4 h-4" />
                        Generate PO Draft PDF
                      </button>
                      {canShowConvertPRToPO(selectedRequest, user) && (
                        <button 
                          onClick={() => openConvertPoModal(selectedRequest)}
                          className="flex items-center gap-2 px-4 py-2 bg-indigo-600 rounded-xl text-sm font-bold text-white hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-100"
                        >
                          <RefreshCw className="w-4 h-4" />
                          Convert PR → PO
                        </button>
                      )}
                    </>
                  )}
                  {canShowPurchasingFinalDecision(selectedRequest, user) && !isEditing && (
                    <>
                      <button
                        onClick={() => openPurchasingDecisionModal(selectedRequest, 'rejected')}
                        className="flex items-center gap-2 px-4 py-2 bg-red-50 border border-red-200 rounded-xl text-sm font-bold text-red-700 hover:bg-red-100 transition-colors"
                      >
                        <XCircle className="w-4 h-4" />
                        Reject PR (Resubmittable)
                      </button>
                      <button
                        onClick={() => openPurchasingDecisionModal(selectedRequest, 'cancelled')}
                        className="flex items-center gap-2 px-4 py-2 bg-zinc-200 border border-zinc-300 rounded-xl text-sm font-bold text-zinc-800 hover:bg-zinc-300 transition-colors"
                        title="Cancellation is final and cannot be resubmitted"
                      >
                        <X className="w-4 h-4" />
                        Cancel PR (Final)
                      </button>
                    </>
                  )}
                  {canRequesterCancelSelected && !isEditing && (
                    <button
                      onClick={() => handleRequesterCancel(selectedRequest.id)}
                      disabled={loading}
                      className="flex items-center gap-2 px-4 py-2 bg-zinc-200 border border-zinc-300 rounded-xl text-sm font-bold text-zinc-800 hover:bg-zinc-300 transition-colors disabled:opacity-50"
                      title="Cancel before any approver has approved (final action)"
                    >
                      <X className="w-4 h-4" />
                      Cancel Request
                    </button>
                  )}
                  <button 
                    onClick={() => handlePrintPR(selectedRequest)} 
                    className="flex items-center gap-2 px-4 py-2 bg-white border border-zinc-200 rounded-xl text-sm font-bold text-zinc-600 hover:bg-zinc-50 transition-colors"
                    title={isPR_Only(selectedRequest) ? 'View PR form' : isSR_Only(selectedRequest) ? 'View SR form' : 'View PDF'}
                  >
                    <FileText className="w-4 h-4" />
                    {isPR_Only(selectedRequest) ? 'View PR form' : isSR_Only(selectedRequest) ? 'View SR form' : 'View PDF'}
                  </button>
                  <button onClick={() => {
                    setSelectedRequest(null);
                    setIsEditing(false);
                    setDetailsViewMode('details');
                    setDetailsPdfPreview((prev) => {
                      if (prev?.url.startsWith('blob:')) URL.revokeObjectURL(prev.url);
                      return null;
                    });
                  }} className="text-zinc-400 hover:text-zinc-600">
                    <XCircle className="w-6 h-6" />
                  </button>
                </div>
              </div>
              
              <div className="p-0 overflow-y-auto flex-1 bg-zinc-50/50">
                {isEditing ? (
                  <div className="p-8 w-full">
                    <div className="bg-white rounded-2xl shadow-sm border border-zinc-200 p-8 space-y-8">
                      <div className="flex items-center gap-3 mb-6">
                        <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center">
                          <Edit2 className="w-5 h-5 text-indigo-600" />
                        </div>
                        <div>
                          <h3 className="text-xl font-bold text-zinc-900">Edit Request</h3>
                          <p className="text-sm text-zinc-500">Update your request details and line items</p>
                        </div>
                      </div>

                      <div className="space-y-6">
                        {isProcurementPRorPORequest(selectedRequest) && (
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Currency</label>
                              <select
                                className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                value={editData.currency}
                                onChange={(e) => setEditData({ ...editData, currency: e.target.value })}
                                required
                              >
                                <option value="">Select currency…</option>
                                {CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">{procurementTaxRateFormLabel(selectedRequest.entity)}</label>
                              <div className="flex items-center gap-3">
                                <input
                                  type="number"
                                  step="0.01"
                                  min="0"
                                  max="100"
                                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                  value={editData.tax_rate}
                                  onChange={(e) => setEditData({ ...editData, tax_rate: parseFloat(e.target.value) || 0 })}
                                />
                                <span className="text-xs text-zinc-500 font-bold uppercase tracking-tighter">%</span>
                              </div>
                            </div>
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Discount (%)</label>
                              <div className="flex items-center gap-3">
                                <input
                                  type="number"
                                  step="0.01"
                                  min="0"
                                  max="100"
                                  className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                  value={editData.discount_rate !== undefined ? editData.discount_rate : 0}
                                  onChange={(e) => setEditData({ ...editData, discount_rate: parseFloat(e.target.value) || 0 })}
                                />
                                <span className="text-xs text-zinc-500 font-bold uppercase tracking-tighter">%</span>
                              </div>
                            </div>
                          </div>
                        )}
                        {isSRRequest(selectedRequest) && (
                          <div className="max-w-md">
                            <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Section</label>
                            <select
                              className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                              value={sectionSelectionFromStored(editData.section)}
                              onChange={(e) => setEditData({ ...editData, section: e.target.value })}
                              required
                            >
                              <option value="" disabled>Select section...</option>
                              {SECTION_OPTIONS.map((opt) => (
                                <option key={opt} value={opt}>{opt}</option>
                              ))}
                            </select>
                          </div>
                        )}
                        {isPRRequest(selectedRequest) && (
                          <div className="max-w-2xl">
                            <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Suggested supplier</label>
                            <input
                              type="text"
                              className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                              value={editData.suggested_supplier ?? ''}
                              onChange={(e) => setEditData({ ...editData, suggested_supplier: e.target.value })}
                              required
                            />
                            <p className="text-xs text-zinc-500 mt-1">One supplier for the entire PR.</p>
                          </div>
                        )}
                        <div>
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Request Title</label>
                          <input
                            type="text"
                            className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                            value={editData.title}
                            onChange={(e) => setEditData({ ...editData, title: e.target.value })}
                          />
                        </div>
                        {isPO_Only(selectedRequest) && isPurchasing && isWorkflowRequestPending(selectedRequest) && (
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 max-w-3xl">
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Requester login</label>
                              <input
                                type="text"
                                className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                autoComplete="off"
                                placeholder="Username in system"
                                value={editData.requester_username}
                                onChange={(e) => setEditData({ ...editData, requester_username: e.target.value })}
                              />
                            </div>
                            <div>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Requester name</label>
                              <input
                                type="text"
                                className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500"
                                placeholder="Name on document"
                                value={editData.requester_name}
                                onChange={(e) => setEditData({ ...editData, requester_name: e.target.value })}
                              />
                            </div>
                            <p className="text-xs text-zinc-500 sm:col-span-2">
                              Changing login reassigns the requester and sets department from their profile. Changing only the name updates department from the profile of the current requester.
                            </p>
                          </div>
                        )}
                        {!isProcurementPRorPORequest(selectedRequest) && (
                        <div>
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 block">Details</label>
                          <textarea
                            className="w-full px-4 py-3 rounded-xl border border-zinc-200 text-sm outline-none focus:ring-2 focus:ring-indigo-500 min-h-[120px]"
                            value={editData.details}
                            onChange={(e) => setEditData({ ...editData, details: e.target.value })}
                          />
                        </div>
                        )}

                        <div>
                          <div className="flex items-center justify-between mb-4">
                            <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">
                              {isPRRequest(selectedRequest) ? 'PR line items' : isSRRequest(selectedRequest) ? 'Stock requisition line items' : isPO_Only(selectedRequest) ? 'PO line items' : 'Line items'}
                            </label>
                            <button
                              type="button"
                              onClick={() => {
                                const newItems = [...editData.line_items];
                                const emptyItem: any = { [LINE_ITEM_REMARKS_KEY]: '' };
                                procurementGridColumns(selectedRequest).forEach(col => emptyItem[col] = '');
                                newItems.push(emptyItem);
                                setEditData({ ...editData, line_items: newItems });
                              }}
                              className="flex items-center gap-1 text-[10px] font-bold text-indigo-600 hover:text-indigo-700 uppercase tracking-widest"
                            >
                              <Plus className="w-3 h-3" />
                              Add Item
                            </button>
                          </div>
                          <div className="border border-zinc-200 rounded-xl overflow-hidden">
                            <table className="w-full text-sm">
                              <thead className="bg-zinc-50 border-b border-zinc-200">
                                <tr>
                                  {procurementGridColumns(selectedRequest).map(col => (
                                    <th key={col} className="px-4 py-2 font-bold text-zinc-600 text-left">{col}</th>
                                  ))}
                                  <th className="w-10"></th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-zinc-200">
                                {editData.line_items.map((item, idx) => {
                                  const cols = procurementGridColumns(selectedRequest);
                                  return (
                                  <React.Fragment key={idx}>
                                  <tr>
                                    {cols.map(col => (
                                      <td key={col} className="px-2 py-2">
                                        {isCostCenterGridColumn(col) ? (
                                          <SearchableSelect
                                            options={costCenterOptions}
                                            value={item[col] || ''}
                                            onChange={(val) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = { ...newItems[idx], [col]: val };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                            placeholder="Select Cost Center..."
                                            loading={costCentersLoading}
                                            loadFailed={costCentersLoadFailed}
                                            loadErrorHint={costCentersFetchError}
                                          />
                                        ) : (isSRRequest(selectedRequest) && isSpareLocationColumn(col)) ? (
                                          <SearchableSelect
                                            options={costCenterOptions}
                                            value={item[col] || ''}
                                            onChange={(val) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = { ...newItems[idx], [col]: val };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                            placeholder="Select Spare for (Location)..."
                                            loading={costCentersLoading}
                                            loadFailed={costCentersLoadFailed}
                                            loadErrorHint={costCentersFetchError}
                                          />
                                        ) : isProcurementLineItemDateColumn(col) ? (
                                          <input
                                            type="date"
                                            className="w-full px-2 py-1.5 rounded border border-zinc-100 text-xs focus:ring-1 focus:ring-indigo-500 outline-none"
                                            value={htmlDateValueFromStored(item[col])}
                                            onChange={(e) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = { ...newItems[idx], [col]: e.target.value };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                          />
                                        ) : (
                                          <input
                                            type={isProcurementNumericGridColumn(col) ? 'number' : 'text'}
                                            className="w-full px-2 py-1.5 rounded border border-zinc-100 text-xs focus:ring-1 focus:ring-indigo-500 outline-none"
                                            value={col === REMARKS_LINE_COL ? (lineItemRemarksDisplay(item) || '') : (item[col] || '')}
                                            onChange={(e) => {
                                              const newItems = [...editData.line_items];
                                              newItems[idx] = col === REMARKS_LINE_COL
                                                ? mergeLineItemRemarksWrite(newItems[idx], e.target.value)
                                                : { ...newItems[idx], [col]: e.target.value };
                                              setEditData({ ...editData, line_items: newItems });
                                            }}
                                          />
                                        )}
                                      </td>
                                    ))}
                                    <td className="px-2 py-2">
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const newItems = editData.line_items.filter((_, i) => i !== idx);
                                          setEditData({ ...editData, line_items: newItems });
                                        }}
                                        className="text-red-400 hover:text-red-600"
                                      >
                                        <Trash2 className="w-4 h-4" />
                                      </button>
                                    </td>
                                  </tr>
                                  <tr className="bg-zinc-50/90">
                                    <td colSpan={cols.length + 1} className="px-4 py-2">
                                      <label className="text-[10px] font-bold text-zinc-500 uppercase">Line remarks (optional)</label>
                                      <textarea
                                        rows={2}
                                        className="w-full mt-1 px-2 py-1.5 rounded border border-zinc-200 text-xs focus:ring-1 focus:ring-indigo-500 outline-none resize-y min-h-[52px]"
                                        value={item[LINE_ITEM_REMARKS_KEY] ?? ''}
                                        onChange={(e) => {
                                          const newItems = [...editData.line_items];
                                          newItems[idx] = { ...newItems[idx], [LINE_ITEM_REMARKS_KEY]: e.target.value };
                                          setEditData({ ...editData, line_items: newItems });
                                        }}
                                        placeholder="Extra notes for this line…"
                                      />
                                    </td>
                                  </tr>
                                  </React.Fragment>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>

                        <div className="space-y-3">
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block">Attachments</label>
                          <div className="flex flex-wrap gap-2">
                            {editAttachmentKeep.map((att) => (
                              <div key={`keep-${att.id}`} className="flex items-center gap-2 bg-zinc-100 px-3 py-1 rounded-full text-xs text-zinc-600 border border-zinc-200">
                                <Paperclip className="w-3 h-3" />
                                <span>{att.file_name}</span>
                                <button
                                  type="button"
                                  onClick={() => setEditAttachmentKeep((prev) => prev.filter((a) => a.id !== att.id))}
                                  title="Remove existing attachment"
                                >
                                  <Trash2 className="w-3 h-3 hover:text-red-500" />
                                </button>
                              </div>
                            ))}
                            {editAttachmentAdd.map((att) => (
                              <div key={att.id} className="flex items-center gap-2 bg-indigo-50 px-3 py-1 rounded-full text-xs text-indigo-700 border border-indigo-100">
                                <Upload className="w-3 h-3" />
                                <span>{att.name}</span>
                                <button
                                  type="button"
                                  onClick={() => setEditAttachmentAdd((prev) => prev.filter((a) => a.id !== att.id))}
                                  title="Remove new attachment"
                                >
                                  <Trash2 className="w-3 h-3 hover:text-red-500" />
                                </button>
                              </div>
                            ))}
                          </div>
                          <label className="flex flex-col items-center justify-center w-full h-20 border-2 border-dashed border-zinc-300 rounded-xl cursor-pointer hover:bg-zinc-50 transition-colors">
                            <div className="flex flex-col items-center justify-center">
                              <Upload className="w-5 h-5 text-zinc-400 mb-1" />
                              <p className="text-xs text-zinc-500">Re-upload or add more files</p>
                            </div>
                            <input type="file" className="hidden" multiple onChange={handleEditAttachmentAdd} />
                          </label>
                        </div>

                      <div className="flex justify-end gap-3 pt-6 border-t border-zinc-100">
                        <button
                          onClick={() => setIsEditing(false)}
                          className="px-6 py-2 rounded-xl border border-zinc-200 text-sm font-bold text-zinc-600 hover:bg-zinc-50 transition-colors"
                        >
                          Cancel
                        </button>
                        <button
                          onClick={handleSaveEdit}
                          disabled={loading}
                          className="px-8 py-2 rounded-xl bg-indigo-600 text-white text-sm font-bold hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-200 disabled:opacity-50"
                        >
                          {loading ? 'Saving...' : 'Save Changes'}
                        </button>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="w-full bg-white my-8 shadow-sm border border-zinc-200 rounded-xl overflow-hidden">
                  {/* Document Header */}
                  <div className="p-8 border-b border-zinc-100 bg-zinc-50/30">
                    <div className="flex justify-between items-start mb-8">
                      <div>
                        <h1 className="text-2xl font-black text-zinc-900 uppercase tracking-tight">
                          {isPRRequest(selectedRequest)
                              ? 'Purchase Requisition Form'
                              : isSRRequest(selectedRequest)
                                ? 'Stock Item Requisition Form'
                                : isPO_Only(selectedRequest)
                                  ? 'Purchase Order'
                                  : 'Workflow Request'}
                        </h1>
                        <p className="text-sm text-zinc-500 font-medium">Ref: #{displayRequestSerial(selectedRequest)}</p>
                      </div>
                      <div className="text-right">
                        <span className={cn(
                          "px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider",
                          workflowRequestStatusBadgeClass(selectedRequest)
                        )}>
                          {formatWorkflowRequestStatusLabel(selectedRequest)}
                        </span>
                        <p className="text-[10px] text-zinc-400 mt-2 font-bold uppercase tracking-widest">
                          {formatDateMYT(selectedRequest.created_at)}
                        </p>
                      </div>
                    </div>

                    {/* Summary Cards */}
                    <div className="grid grid-cols-4 gap-4 mb-8">
                      <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                        <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Total Amount</p>
                        <p className="text-xl font-black text-indigo-600">
                          {procurementCurrencyPrefix(selectedRequest.currency)}
                          {formatProcurementMoney(
                            isProcurementPRorPORequest(selectedRequest)
                              ? procurementMoneyTotals(selectedRequest).total
                              : ((selectedRequest.line_items?.reduce((sum, item) => sum + (parseFloat(item['Quantity'] || '0') * parseFloat(item['Unit Price'] || '0')), 0) || 0) *
                                  (1 + (selectedRequest.tax_rate !== undefined ? selectedRequest.tax_rate : 0.18)))
                          )}
                        </p>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                        <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Items Count</p>
                        <p className="text-xl font-black text-zinc-900">{selectedRequest.line_items?.length || 0} Items</p>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                        <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Requester</p>
                        <p className="text-sm font-bold text-zinc-900 truncate">{selectedRequest.requester_name}</p>
                        <p className="text-[10px] text-zinc-400 font-bold uppercase">{selectedRequest.department}</p>
                      </div>
                      <div className="bg-white p-4 rounded-xl border border-zinc-100 shadow-sm">
                        <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-1">Current Step</p>
                        <p className="text-sm font-bold text-zinc-900">
                          {isWorkflowRequestPending(selectedRequest) 
                            ? selectedRequest.template_steps[selectedRequest.current_step_index]?.label || 'Processing'
                            : formatWorkflowRequestStatusLabel(selectedRequest)}
                        </p>
                        <p className="text-[10px] text-zinc-400 font-bold uppercase">Workflow Progress</p>
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-8">
                      <div>
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Requester Information</label>
                        <p className="font-bold text-zinc-900">{selectedRequest.requester_name}</p>
                        <p className="text-sm text-zinc-500">{selectedRequest.department}</p>
                      </div>
                      <div className="text-right">
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Template Type</label>
                        <p className="font-bold text-zinc-900">{selectedRequest.template_name}</p>
                        <p className="text-sm text-zinc-500">Category: {selectedRequest.category || 'General'}</p>
                      </div>
                    </div>

                    {requestHasChosenApprover(selectedRequest) ? (
                      <div className="mt-6 pt-6 border-t border-zinc-100">
                        <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Chosen approver</label>
                        <p className="font-bold text-zinc-900">{chosenApproverNameLabel(selectedRequest)}</p>
                        {String(selectedRequest.assigned_approver_designation ?? '').trim() ? (
                          <p className="text-sm text-zinc-500">{String(selectedRequest.assigned_approver_designation).trim()}</p>
                        ) : null}
                      </div>
                    ) : null}

                    {isProcurementPRorPORequest(selectedRequest) && (
                      <div className="grid grid-cols-2 gap-8 mt-8 pt-8 border-t border-zinc-100">
                        <div>
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Currency</label>
                          <p className="font-bold text-zinc-900">{procurementCurrencyLabel(selectedRequest.currency)}</p>
                        </div>
                        <div className="text-right">
                          <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mb-1">Cost Center</label>
                          <p className="font-bold text-zinc-900">
                            {String(selectedRequest.cost_center || '').trim()
                              ? formatCostCenterCatalogCell(selectedRequest.cost_center, costCenterOptions)
                              : 'Not Specified'}
                          </p>
                          {isSRRequest(selectedRequest) && (
                            <>
                              <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest block mt-3 mb-1">Section</label>
                              <p className="font-bold text-zinc-900">{selectedRequest.section || 'Not Specified'}</p>
                            </>
                          )}
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="p-8 space-y-10">
                    <section>
                      <div className="flex items-center justify-between gap-3 mb-3 border-b border-zinc-100 pb-2">
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">Details</h4>
                        <div className="inline-flex rounded-lg border border-zinc-200 overflow-hidden">
                          <button
                            type="button"
                            onClick={() => setDetailsViewMode('details')}
                            className={cn(
                              "px-3 py-1.5 text-[10px] font-black uppercase tracking-widest transition-colors",
                              detailsViewMode === 'details'
                                ? "bg-indigo-600 text-white"
                                : "bg-white text-zinc-500 hover:bg-zinc-50"
                            )}
                          >
                            Details
                          </button>
                          <button
                            type="button"
                            onClick={() => handleShowRequestPdfInline(selectedRequest)}
                            className={cn(
                              "px-3 py-1.5 text-[10px] font-black uppercase tracking-widest transition-colors border-l border-zinc-200 inline-flex items-center gap-1.5",
                              detailsViewMode === 'pdf'
                                ? "bg-indigo-600 text-white"
                                : "bg-white text-zinc-500 hover:bg-zinc-50"
                            )}
                          >
                            <FileText className="w-3 h-3" />
                            PDF
                          </button>
                        </div>
                      </div>
                      {detailsViewMode === 'pdf' ? (
                        <div className="w-full h-[72vh] border border-zinc-200 rounded-xl overflow-hidden bg-zinc-100">
                          {detailsPdfPreview ? (
                            <iframe
                              src={`${detailsPdfPreview.url}#view=FitH`}
                              className="w-full h-full border-none"
                              title="Request PDF Preview"
                            />
                          ) : (
                            <div className="h-full flex items-center justify-center text-sm text-zinc-500">
                              PDF is not ready.
                            </div>
                          )}
                        </div>
                      ) : (
                        !isProcurementPRorPORequest(selectedRequest) ? (
                          <p className="text-zinc-700 whitespace-pre-wrap leading-relaxed">{selectedRequest.details}</p>
                        ) : (
                          <p className="text-zinc-500 text-sm">Switch to PDF to view the full form layout.</p>
                        )
                      )}
                    </section>

                    {isPRRequest(selectedRequest) && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-2 border-b border-zinc-100 pb-2">
                          Suggested supplier
                        </h4>
                        <p className="text-sm text-zinc-800 font-medium">
                          {prSuggestedSupplierDisplay(selectedRequest) || '—'}
                        </p>
                      </section>
                    )}

                    {/* Line Items */}
                    {selectedRequest.line_items && selectedRequest.line_items.length > 0 && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-3 border-b border-zinc-100 pb-2">
                          {isPRRequest(selectedRequest) ? 'PR line items' : isSRRequest(selectedRequest) ? 'Stock requisition line items' : isPO_Only(selectedRequest) ? 'PO line items' : 'Line items'}
                        </h4>
                        <div className="border border-zinc-200 rounded-lg overflow-hidden">
                          <table className="w-full text-sm">
                            <thead className="bg-zinc-50 border-b border-zinc-200">
                              <tr>
                                <th className="px-4 py-2 font-bold text-zinc-600 w-12 text-center">No</th>
                                {procurementGridColumns(selectedRequest).map(col => (
                                  <th key={col} className="px-4 py-2 font-bold text-zinc-600 text-left">{col}</th>
                                ))}
                                {procurementRowShowsLineTotal(selectedRequest) && <th className="px-4 py-2 font-bold text-zinc-600 text-right">Total</th>}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-zinc-200">
                              {selectedRequest.line_items.map((item, idx) => {
                                const cols = procurementGridColumns(selectedRequest);
                                const showLineTotal = procurementRowShowsLineTotal(selectedRequest);
                                const qty = lineItemQtyForDisplay(item);
                                const price = parseFloat(item['Unit Price'] || '0');
                                const total = qty * price;
                                const lineRem = item[LINE_ITEM_REMARKS_KEY] && String(item[LINE_ITEM_REMARKS_KEY]).trim();
                                return (
                                  <React.Fragment key={item.id || idx}>
                                  <tr className={idx % 2 === 0 ? 'bg-white' : 'bg-zinc-50/30'}>
                                    <td className="px-4 py-3 text-center text-zinc-500 font-medium">{idx + 1}</td>
                                    {cols.map(col => (
                                      <td key={col} className="px-4 py-3 text-zinc-700">
                                        {(col === 'Unit Price' || col === 'Price' || col === 'Amount')
                                          ? `${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(Number(item[col] || 0))}`.trim()
                                          : (col === REMARKS_LINE_COL
                                            ? (lineItemRemarksDisplay(item) || '-')
                                            : (isCostCenterGridColumn(col) || (isSRRequest(selectedRequest) && isSpareLocationColumn(col)))
                                              ? formatCostCenterCatalogCell(item[col], costCenterOptions)
                                              : (item[col] || '-'))}
                                      </td>
                                    ))}
                                    {showLineTotal && (
                                      <td className="px-4 py-3 text-right font-bold text-zinc-900">
                                        {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(total)}`.trim()}
                                      </td>
                                    )}
                                  </tr>
                                  {lineRem && (
                                    <tr className="bg-indigo-50/40 border-t border-indigo-100/80">
                                      <td colSpan={cols.length + 1 + (showLineTotal ? 1 : 0)} className="px-4 py-2.5 text-xs text-zinc-700">
                                        <span className="font-bold text-zinc-500 uppercase tracking-wide text-[10px]">Line remarks: </span>
                                        <span className="whitespace-pre-wrap">{lineRem}</span>
                                      </td>
                                    </tr>
                                  )}
                                  </React.Fragment>
                                );
                              })}
                            </tbody>
                            {procurementRowShowsLineTotal(selectedRequest) && (() => {
                              const m = procurementMoneyTotals(selectedRequest);
                              const gc = procurementGridColumns(selectedRequest);
                              return (
                              <tfoot className="bg-zinc-50/50 font-bold">
                                <tr>
                                  <td colSpan={gc.length + 1} className="px-4 py-2 text-right text-zinc-500">Subtotal</td>
                                  <td className="px-4 py-2 text-right text-zinc-900">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.subtotal)}`.trim()}
                                  </td>
                                </tr>
                                <tr>
                                  <td colSpan={gc.length + 1} className="px-4 py-2 text-right text-zinc-500">Discount ({(m.discountRate * 100).toFixed(0)}%)</td>
                                  <td className="px-4 py-2 text-right text-zinc-900">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.discountAmount)}`.trim()}
                                  </td>
                                </tr>
                                <tr>
                                  <td colSpan={gc.length + 1} className="px-4 py-2 text-right text-zinc-500">{procurementTaxLabelForEntity(selectedRequest.entity)} ({(m.taxRate * 100).toFixed(0)}%)</td>
                                  <td className="px-4 py-2 text-right text-zinc-900">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.taxAmount)}`.trim()}
                                  </td>
                                </tr>
                                <tr className="bg-zinc-100/50 text-lg">
                                  <td colSpan={gc.length + 1} className="px-4 py-4 text-right text-zinc-900 uppercase tracking-tight">Total Amount</td>
                                  <td className="px-4 py-4 text-right text-indigo-600 font-black">
                                    {`${procurementCurrencyPrefix(selectedRequest.currency)}${formatProcurementMoney(m.total)}`.trim()}
                                  </td>
                                </tr>
                              </tfoot>
                              );
                            })()}
                          </table>
                        </div>
                      </section>
                    )}

                    {isPO_Only(selectedRequest) && isPurchasing && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-3 border-b border-zinc-100 pb-2">
                          Real Purchase Order
                        </h4>
                        <label className="inline-flex items-center gap-2 px-3 py-2 rounded-lg border border-zinc-200 bg-zinc-50 hover:bg-zinc-100 cursor-pointer text-xs font-semibold text-zinc-700 transition-colors">
                          <Upload className="w-4 h-4" />
                          {loading ? 'Uploading…' : 'Upload PO file (optional)'}
                          <input
                            type="file"
                            className="hidden"
                            disabled={loading}
                            onChange={(e) => {
                              const f = e.target.files?.[0];
                              if (!f) return;
                              void handleUploadRealPoAttachment(f);
                              e.currentTarget.value = '';
                            }}
                          />
                        </label>
                        <p className="text-[11px] text-zinc-500 mt-2">
                          You can upload the official PO now or later.
                        </p>
                      </section>
                    )}

                    {/* Attachments */}
                    {attachments.length > 0 && (
                      <section>
                        <h4 className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-3 border-b border-zinc-100 pb-2">Attachments</h4>
                        <div className="grid grid-cols-2 gap-3">
                          {attachments.map((att) => (
                            <div
                              key={att.id}
                              onClick={() => {
                                void openWorkflowRequestAttachment(selectedRequest.id, att, setViewingPdf);
                              }}
                              className="flex items-center gap-3 p-3 bg-zinc-50 rounded-xl border border-zinc-200 hover:border-indigo-300 transition-all group cursor-pointer"
                            >
                              <div className="w-8 h-8 bg-white rounded-lg flex items-center justify-center border border-zinc-100 group-hover:bg-indigo-50 transition-colors">
                                <FileText className="w-4 h-4 text-indigo-500" />
                              </div>
                              <div className="flex-1 min-w-0">
                                <p className="text-xs font-bold text-zinc-700 truncate">{att.file_name}</p>
                                <p className="text-[10px] text-zinc-400 uppercase font-bold">
                                  {(att.file_type === 'application/pdf' || att.file_name.toLowerCase().endsWith('.pdf')) ? 'View PDF' : 'Download File'}
                                </p>
                              </div>
                              {(att.file_type === 'application/pdf' || att.file_name.toLowerCase().endsWith('.pdf')) ? (
                                <Eye className="w-3 h-3 text-zinc-300 group-hover:text-indigo-500" />
                              ) : (
                                <Download className="w-3 h-3 text-zinc-300 group-hover:text-indigo-500" />
                              )}
                            </div>
                          ))}
                        </div>
                      </section>
                    )}

                    {/* Signatures & Approvals */}
                    <section className="pt-8 border-t-2 border-dashed border-zinc-100">
                      <div className="grid grid-cols-2 gap-12">
                        {/* Requester Signature */}
                        <div>
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-4">Requested By</p>
                          {hasRequesterSignatureProof(selectedRequest) ? (
                            <div className="space-y-2">
                              {isSignatureImageDataUrl(selectedRequest.requester_signature) ? (
                                <img src={selectedRequest.requester_signature!} alt="Requester Signature" className="h-16 object-contain" referrerPolicy="no-referrer" />
                              ) : (
                                <div className="rounded-lg border border-emerald-200 bg-emerald-50/80 px-3 py-2">
                                  <p className="text-xs font-semibold text-emerald-800">Electronic signature</p>
                                  {selectedRequest.requester_signed_at && (
                                    <p className="text-[10px] text-emerald-700/90 mt-0.5">{formatSignatureProofTimestamp(selectedRequest.requester_signed_at)}</p>
                                  )}
                                </div>
                              )}
                              <div className="pt-2 border-t border-zinc-200">
                                <p className="text-sm font-bold text-zinc-900">{selectedRequest.requester_name}</p>
                                <p className="text-xs text-zinc-500">{selectedRequest.department}</p>
                                <p className="text-[10px] text-zinc-400 mt-1">{formatDateMYT(selectedRequest.created_at)}</p>
                              </div>
                            </div>
                          ) : (
                            <div className="h-24 border-2 border-dashed border-zinc-100 rounded-xl flex items-center justify-center">
                              <p className="text-xs text-zinc-300 italic">No signature</p>
                            </div>
                          )}
                        </div>

                        {/* Approvals */}
                        <div className="space-y-8">
                          <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mb-4">Approval Workflow</p>
                          {selectedRequest.template_steps.map((step, i) => {
                            const approval = [...approvals].reverse().find((a) => a.step_index === i);
                            const isCurrent = isWorkflowRequestPending(selectedRequest) && selectedRequest.current_step_index === i;
                            
                            return (
                              <div key={step.id} className="relative pl-8 border-l-2 border-zinc-100 pb-8 last:pb-0">
                                <div className={cn(
                                  "absolute -left-[9px] top-0 w-4 h-4 rounded-full border-2 border-white",
                                  approval ? (approval.status === 'approved' ? "bg-emerald-500" : "bg-red-500") :
                                  isCurrent ? "bg-amber-500 animate-pulse" : "bg-zinc-200"
                                )} />
                                
                                <div className="space-y-2">
                                  <div className="flex justify-between items-start">
                                    <div>
                                      <p className="text-xs font-black text-zinc-900 uppercase tracking-tight">{step.label}</p>
                                      <p className="text-[10px] text-zinc-400 font-bold uppercase">Role: {step.approverRole}</p>
                                    </div>
                                    {approval && (
                                      <span className={cn(
                                        "text-[10px] font-black uppercase px-2 py-0.5 rounded",
                                        approval.status === 'approved' ? "bg-emerald-100 text-emerald-700" : "bg-red-100 text-red-700"
                                      )}>
                                        {approval.status}
                                      </span>
                                    )}
                                  </div>

                                  {approval && (
                                    <div className="mt-4 p-4 bg-zinc-50 rounded-xl border border-zinc-100">
                                      {approval.comment && <p className="text-xs text-zinc-600 italic mb-3">"{approval.comment}"</p>}
                                      <div className="flex items-center gap-3">
                                        {isSignatureImageDataUrl(approval.approver_signature) ? (
                                          <img src={approval.approver_signature!} alt="Approver Signature" className="h-10 object-contain" referrerPolicy="no-referrer" />
                                        ) : approval.status.toLowerCase() === 'approved' ? (
                                          <div className="shrink-0 rounded-lg border border-emerald-200 bg-emerald-50/80 px-2 py-1.5">
                                            <p className="text-[10px] font-semibold text-emerald-800">E-signed</p>
                                            <p className="text-[9px] text-emerald-700/90">{formatSignatureProofTimestamp(approval.created_at)}</p>
                                          </div>
                                        ) : null}
                                        <div>
                                          <p className="text-xs font-bold text-zinc-900">{approval.approver_name}</p>
                                          {approval.signed_by_name ? (
                                            <p className="text-[10px] text-indigo-600 font-medium">
                                              Signed by proxy: {approval.signed_by_name}
                                            </p>
                                          ) : null}
                                          <p className="text-[10px] text-zinc-400">{formatDateMYT(approval.created_at)}</p>
                                        </div>
                                      </div>
                                    </div>
                                  )}

                                  {isCurrent && (
                                    <div className="mt-4 p-4 bg-amber-50 border border-amber-200 rounded-xl">
                                      <p className="text-xs font-bold text-amber-700 flex items-center gap-2">
                                        <Clock className="w-3 h-3" />
                                        Awaiting decision...
                                      </p>
                                    </div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </section>

                    {isWorkflowRequestRejected(selectedRequest) && user.id === selectedRequest.requester_id && (
                      <section className="pt-12 border-t-2 border-zinc-100">
                        <div className="bg-rose-50 p-8 rounded-2xl border border-rose-200">
                          <div className="flex items-center gap-3 mb-6">
                            <div className="w-10 h-10 bg-rose-600 rounded-xl flex items-center justify-center shadow-lg shadow-rose-200">
                              <RefreshCw className="w-5 h-5 text-white" />
                            </div>
                            <div>
                              <h3 className="text-lg font-black text-zinc-900 uppercase tracking-tight">Resubmit Request</h3>
                              <p className="text-xs text-zinc-500">This request was rejected. You can resubmit it to restart the workflow.</p>
                            </div>
                          </div>
                          {(() => {
                            const ra = purchasingRejectApprovalForRequest({ ...selectedRequest, approvals });
                            const reason = String(ra?.comment || '').trim();
                            const who = String(ra?.approver_name || '').trim();
                            const when = ra?.created_at ? formatDateTimeMYT(ra.created_at) : '';
                            if (!reason && !who && !when) return null;
                            return (
                              <div className="mb-6 bg-rose-100 border border-rose-300 rounded-xl p-4">
                                <div className="text-[10px] font-black text-rose-900 uppercase tracking-widest mb-1">
                                  Rejected by purchasing
                                </div>
                                <div className="text-xs text-rose-900">
                                  {who ? <span className="font-bold">{who}</span> : null}
                                  {when ? <span className="text-rose-800/80">{who ? ` • ${when}` : when}</span> : null}
                                </div>
                                {reason ? (
                                  <div className="mt-2 text-sm text-rose-950 whitespace-pre-wrap break-words">“{reason}”</div>
                                ) : null}
                              </div>
                            );
                          })()}
                          <button
                            onClick={() => handleResubmit(selectedRequest.id)}
                            disabled={loading}
                            className="w-full bg-indigo-600 text-white py-4 rounded-xl font-black uppercase tracking-widest hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 flex items-center justify-center gap-2 disabled:opacity-50"
                          >
                            <Send className="w-5 h-5" />
                            Resubmit Now
                          </button>
                        </div>
                      </section>
                    )}
                    {isWorkflowRequestCancelled(selectedRequest) && (
                      <section className="pt-12 border-t-2 border-zinc-100">
                        <div className="bg-rose-50 p-6 rounded-2xl border border-rose-200">
                          {(() => {
                            const requesterCancel = requesterCancelApprovalForRequest({ ...selectedRequest, approvals });
                            const purchasingCancel = purchasingCancelApprovalForRequest({ ...selectedRequest, approvals });
                            const ca = requesterCancel || purchasingCancel;
                            const reason = String(ca?.comment || '').trim();
                            const who = String(ca?.approver_name || '').trim();
                            const when = ca?.created_at ? formatDateTimeMYT(ca.created_at) : '';
                            const byRequester = !!requesterCancel;
                            return (
                              <div className="space-y-2">
                                <p className="text-sm font-bold text-rose-800">
                                  {byRequester ? 'Cancelled by requester' : 'Cancelled by purchasing team'}
                                  {who ? ` — ${who}` : ''}{when ? ` (${when})` : ''}. This is final (cannot be resubmitted).
                                </p>
                                {reason ? (
                                  <div className="text-sm text-rose-950 bg-rose-100 border border-rose-300 rounded-xl p-3">
                                    <div className="text-[10px] font-black text-rose-900 uppercase tracking-widest mb-1">Cancellation reason</div>
                                    <div className="whitespace-pre-wrap break-words">{reason}</div>
                                  </div>
                                ) : (
                                  <p className="text-xs text-rose-700/80 italic">No cancellation reason was provided.</p>
                                )}
                              </div>
                            );
                          })()}
                        </div>
                      </section>
                    )}
                  </div>
                </div>
              )}
            </div>
          </motion.div>
          </div>
        )}
        {viewingPdf && (
          <PdfViewer
            url={viewingPdf.url}
            fileName={viewingPdf.fileName}
            onDownload={() => downloadFileFromUrl(viewingPdf.url, viewingPdf.fileName)}
            onClose={() => {
              if (viewingPdf.url.startsWith("blob:")) URL.revokeObjectURL(viewingPdf.url);
              setViewingPdf(null);
            }}
          />
        )}
      </AnimatePresence>
      <ConvertPrToPoModal
        target={convertPoModal}
        loading={loading}
        onClose={() => !loading && setConvertPoModal(null)}
        onConfirm={handleConvertToPOConfirm}
      />
      <PurchasingDecisionModal
        target={purchasingDecisionModal}
        loading={loading}
        onClose={() => !loading && setPurchasingDecisionModal(null)}
        onConfirm={(comment) => {
          if (!purchasingDecisionModal) return;
          handlePurchasingFinalDecision(
            purchasingDecisionModal.id,
            purchasingDecisionModal.decision,
            comment,
            purchasingDecisionModal.entity
          );
        }}
      />
    </div>
  );
};

// --- Main App ---

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [selectedEntity, setSelectedEntity] = useState<string | null>(null);
  const [workflows, setWorkflows] = useState<Workflow[]>([]);
  const [requests, setRequests] = useState<WorkflowRequest[]>([]);
  const [roles, setRoles] = useState<Role[]>([]);
  const [activeTab, setActiveTab] = useState<
    'dashboard' | 'create' | 'templates' | 'submit' | 'requests' | 'admin' | 'procurement' | 'cost-centers'
  >('dashboard');
  const [selectedTemplateForRequest, setSelectedTemplateForRequest] = useState<Workflow | null>(null);
  const [preSelectedRequestId, setPreSelectedRequestId] = useState<number | null>(null);
  const [loading, setLoading] = useState(true);
  const hasMultipleEntities = (user?.entities || []).length > 1;
  const [sidebarCollapsed, setSidebarCollapsed] = useState(() => {
    try {
      return localStorage.getItem(FLOWMASTER_SIDEBAR_COLLAPSED_KEY) === '1';
    } catch {
      return false;
    }
  });

  const toggleSidebarCollapsed = useCallback(() => {
    setSidebarCollapsed((c) => {
      const next = !c;
      try {
        localStorage.setItem(FLOWMASTER_SIDEBAR_COLLAPSED_KEY, next ? '1' : '0');
      } catch {
        /* ignore */
      }
      return next;
    });
  }, []);

  const handleEntityChange = useCallback((nextEntity: string | null) => {
    const ents = (user?.entities || []).map((e) => String(e).trim()).filter(Boolean);
    const normalizedNext = String(nextEntity || '').trim();
    if (!normalizedNext) {
      setSelectedEntity(null);
      api.setActiveEntity(null);
      return;
    }
    if (ents.length > 0 && !ents.includes(normalizedNext)) return;
    setSelectedEntity(normalizedNext);
    api.setActiveEntity(normalizedNext);
  }, [user]);

  const fetchWorkflows = async () => {
    try {
      const data = await api.request('/api/workflows');
      setWorkflows(data);
    } catch (err) {
      console.error(err);
    }
  };

  const fetchRequests = useCallback(async () => {
    try {
      const data = await api.request('/api/workflow-requests', { skipEntity: true });
      setRequests(data);
    } catch (err) {
      console.error(err);
    }
  }, []);

  /** Navigate to Approvals and reload the list (covers sidebar re-clicks while already on this tab). */
  const openRequestsTab = useCallback(() => {
    setSelectedTemplateForRequest(null);
    setActiveTab('requests');
    void fetchRequests();
  }, [fetchRequests]);

  const lastVisibilityRequestsFetchRef = useRef(0);

  const applyEntityForUser = (userData: User) => {
    const ents = userData.entities || [];
    const stored = localStorage.getItem(FLOWMASTER_ENTITY_KEY);
    if (ents.length === 1) {
      handleEntityChange(ents[0]);
      return;
    }
    if (ents.length > 1 && stored && ents.includes(stored)) {
      handleEntityChange(stored);
      return;
    }
    const fallback = ents[0] || null;
    handleEntityChange(fallback);
  };

  const fetchRoles = async () => {
    try {
      const data = await api.request('/api/roles');
      setRoles(data);
    } catch (err) {
      console.error(err);
    }
  };

  const hasPermission = (permission: string) => {
    if (!user) return false;
    if (user.permissions?.includes('admin')) return true;
    return user.permissions?.includes(permission);
  };
  const approverViewRoles = new Set(['approver', 'checker', 'som']);
  const isPreparerUser = (user?.roles || []).some((r) => String(r).toLowerCase() === 'preparer');
  const isApproverViewUser = (user?.roles || []).some((r) => approverViewRoles.has(String(r).toLowerCase()));
  const canSeeDashboard = hasPermission('admin');
  const canSeeTemplateDesigner = hasPermission('create_templates') && !isApproverViewUser && !isPreparerUser;
  const canSeeWorkflowTemplates = (hasPermission('create_templates') || hasPermission('approve_templates')) && !isApproverViewUser && !isPreparerUser;

  const init = async () => {
    const token = localStorage.getItem('token');
    if (token) {
      api.setToken(token);
      try {
        const userData = await api.request('/api/me');
        // Ensure roles is always an array
        if (userData && !userData.roles) {
          userData.roles = userData.role ? [userData.role] : ['user'];
        }
        setUser(userData);
        applyEntityForUser(userData);
        await Promise.all([fetchWorkflows(), fetchRoles()]);
      } catch (err) {
        api.setToken(null);
      }
    }
    setLoading(false);
  };

  useEffect(() => {
    init();
  }, []);

  useEffect(() => {
    const syncEntity = () => {
      const stored = String(localStorage.getItem(FLOWMASTER_ENTITY_KEY) || '').trim() || null;
      const allowed = (user?.entities || []).map((e) => String(e).trim());
      if (stored && allowed.length > 0 && !allowed.includes(stored)) return;
      setSelectedEntity(stored);
    };
    window.addEventListener('storage', syncEntity);
    window.addEventListener(FLOWMASTER_ENTITY_CHANGED_EVENT, syncEntity as EventListener);
    return () => {
      window.removeEventListener('storage', syncEntity);
      window.removeEventListener(FLOWMASTER_ENTITY_CHANGED_EVENT, syncEntity as EventListener);
    };
  }, [user]);

  useEffect(() => {
    if (!user || loading) return;
    const ents = user.entities || [];
    if (ents.length === 0) return;
    void fetchRequests();
  }, [user, selectedEntity, loading, fetchRequests]);

  /** Refetch when returning to the tab/window so multi-tab or background sessions stay aligned. */
  useEffect(() => {
    if (!user || loading) return;
    const onVis = () => {
      if (document.visibilityState !== 'visible') return;
      if (activeTab !== 'requests') return;
      const now = Date.now();
      if (now - lastVisibilityRequestsFetchRef.current < 4000) return;
      lastVisibilityRequestsFetchRef.current = now;
      void fetchRequests();
    };
    document.addEventListener('visibilitychange', onVis);
    return () => document.removeEventListener('visibilitychange', onVis);
  }, [user, loading, activeTab, fetchRequests]);

  useEffect(() => {
    if (!isApproverViewUser && !isPreparerUser && canSeeDashboard) return;
    if (activeTab === 'create' || activeTab === 'templates' || (!canSeeDashboard && activeTab === 'dashboard')) {
      setActiveTab('submit');
      setSelectedTemplateForRequest(null);
    }
  }, [activeTab, isApproverViewUser, isPreparerUser, canSeeDashboard]);

  const handleLogout = () => {
    api.setToken(null);
    setUser(null);
    setSelectedEntity(null);
    toast.success('Logged out');
  };

  if (loading) return <div className="min-h-screen flex items-center justify-center">Loading...</div>;

  if (!user) return <Login onLogin={(u) => { 
    const userData = { ...u };
    if (!userData.roles) {
      userData.roles = (userData as any).role ? [(userData as any).role] : ['user'];
    }
    setUser(userData);
    applyEntityForUser(userData);
    fetchWorkflows();
    fetchRoles();
  }} />;

  if (user.entities && user.entities.length === 0) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-zinc-50 p-4">
        <div className="max-w-md bg-white rounded-2xl border border-zinc-200 p-8 text-center shadow-sm">
          <Building2 className="w-12 h-12 text-amber-500 mx-auto mb-4" />
          <h1 className="text-lg font-bold text-zinc-900 mb-2">No entity access</h1>
          <p className="text-sm text-zinc-600">Your account is not assigned to any entity. Ask an administrator to assign entities before you can use workflows.</p>
          <button
            type="button"
            onClick={handleLogout}
            className="mt-6 w-full bg-zinc-900 text-white py-2 rounded-lg font-semibold hover:bg-zinc-800"
          >
            Log out
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-zinc-50 flex min-w-0">
      <Toaster position="top-right" />
      
      {/* Sidebar — collapsible rail for more horizontal space on laptop/desktop */}
      <aside
        className={cn(
          'bg-white border-r border-zinc-200 flex flex-col shrink-0 transition-[width] duration-200 ease-out overflow-x-hidden overflow-y-auto',
          sidebarCollapsed ? 'w-14' : 'w-64'
        )}
      >
        <div className={cn('border-b border-zinc-100 shrink-0', sidebarCollapsed ? 'p-2' : 'p-4')}>
          <div className={cn('flex items-center gap-2', sidebarCollapsed && 'flex-col')}>
            <div className="w-8 h-8 bg-indigo-600 rounded-lg flex items-center justify-center shadow-lg shadow-indigo-100 shrink-0">
              <Shield className="text-white w-4 h-4" />
            </div>
            {!sidebarCollapsed && (
              <div className="flex-1 min-w-0">
                <span className="font-bold text-lg text-zinc-900 block leading-tight">Approval System</span>
                <p className="text-[10px] text-zinc-400 uppercase font-bold tracking-widest mt-0.5">Approval Platform</p>
              </div>
            )}
            <button
              type="button"
              onClick={toggleSidebarCollapsed}
              title={sidebarCollapsed ? 'Expand navigation' : 'Collapse navigation'}
              aria-label={sidebarCollapsed ? 'Expand navigation' : 'Collapse navigation'}
              className={cn(
                'shrink-0 p-1.5 rounded-lg text-zinc-500 hover:bg-zinc-100 hover:text-zinc-800 transition-colors',
                sidebarCollapsed ? 'mt-1' : 'ml-auto'
              )}
            >
              {sidebarCollapsed ? <ChevronsRight className="w-4 h-4" /> : <ChevronsLeft className="w-4 h-4" />}
            </button>
          </div>
        </div>

        <nav className={cn('flex-1 space-y-1', sidebarCollapsed ? 'p-1.5' : 'p-3')}>
          {canSeeDashboard && (
            <button
              onClick={() => { setActiveTab('dashboard'); setSelectedTemplateForRequest(null); }}
              title="Dashboard"
              aria-label="Dashboard"
              className={cn(
                'w-full flex items-center rounded-xl text-sm font-medium transition-all',
                sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'gap-3 px-3 py-2.5',
                activeTab === 'dashboard' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
              )}
            >
              <LayoutDashboard className="w-4 h-4 shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Dashboard</span>}
            </button>
          )}
          {canSeeTemplateDesigner && (
            <button
              onClick={() => { setActiveTab('create'); setSelectedTemplateForRequest(null); }}
              title="Design Template"
              aria-label="Design Template"
              className={cn(
                'w-full flex items-center rounded-xl text-sm font-medium transition-all',
                sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'gap-3 px-3 py-2.5',
                activeTab === 'create' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
              )}
            >
              <PlusCircle className="w-4 h-4 shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Design Template</span>}
            </button>
          )}
          {canSeeWorkflowTemplates && (
            <div className="relative">
              <button
                onClick={() => { setActiveTab('templates'); setSelectedTemplateForRequest(null); }}
                title="Workflow Templates"
                aria-label="Workflow Templates"
                className={cn(
                  'w-full flex items-center rounded-xl text-sm font-medium transition-all',
                  sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'justify-between gap-2 px-3 py-2.5',
                  activeTab === 'templates' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
                )}
              >
                <div className={cn('flex items-center min-w-0', !sidebarCollapsed && 'gap-3')}>
                  <ClipboardList className="w-4 h-4 shrink-0" />
                  {!sidebarCollapsed && <span className="truncate">Workflow Templates</span>}
                </div>
                {!sidebarCollapsed && workflows.filter((w) => w.status === 'pending').length > 0 && (
                  <span className="bg-amber-500 text-white text-[10px] font-black px-1.5 py-0.5 rounded-full min-w-[18px] text-center shrink-0">
                    {workflows.filter((w) => w.status === 'pending').length}
                  </span>
                )}
              </button>
              {sidebarCollapsed && workflows.filter((w) => w.status === 'pending').length > 0 && (
                <span className="absolute -right-0.5 -top-0.5 bg-amber-500 text-white text-[8px] font-black min-w-[14px] h-[14px] flex items-center justify-center rounded-full">
                  {workflows.filter((w) => w.status === 'pending').length}
                </span>
              )}
            </div>
          )}
          <button
            onClick={() => { setActiveTab('submit'); setSelectedTemplateForRequest(null); }}
            title="Submit Request"
            aria-label="Submit Request"
            className={cn(
              'w-full flex items-center rounded-xl text-sm font-medium transition-all',
              sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'gap-3 px-3 py-2.5',
              activeTab === 'submit' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
            )}
          >
            <Send className="w-4 h-4 shrink-0" />
            {!sidebarCollapsed && <span className="truncate">Submit Request</span>}
          </button>
          <button
            onClick={openRequestsTab}
            title="Approvals & Requests"
            aria-label="Approvals & Requests"
            className={cn(
              'w-full flex items-center rounded-xl text-sm font-medium transition-all',
              sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'gap-3 px-3 py-2.5',
              activeTab === 'requests' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
            )}
          >
            <Inbox className="w-4 h-4 shrink-0" />
            {!sidebarCollapsed && <span className="truncate">Approvals & Requests</span>}
          </button>
          {hasPermission('view_procurement_center') && (
            <button
              onClick={() => { setActiveTab('procurement'); setSelectedTemplateForRequest(null); }}
              title="Procurement Center"
              aria-label="Procurement Center"
              className={cn(
                'w-full flex items-center rounded-xl text-sm font-medium transition-all',
                sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'gap-3 px-3 py-2.5',
                activeTab === 'procurement' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
              )}
            >
              <ShoppingCart className="w-4 h-4 shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Procurement Center</span>}
            </button>
          )}

          {hasPermission('manage_cost_centers') && (
            <button
              onClick={() => { setActiveTab('cost-centers'); setSelectedTemplateForRequest(null); }}
              title="Cost centers"
              aria-label="Cost centers"
              className={cn(
                'w-full flex items-center rounded-xl text-sm font-medium transition-all',
                sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'gap-3 px-3 py-2.5',
                activeTab === 'cost-centers' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
              )}
            >
              <Warehouse className="w-4 h-4 shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Cost centers</span>}
            </button>
          )}

          {hasPermission('manage_users') && (
            <button
              onClick={() => { setActiveTab('admin'); setSelectedTemplateForRequest(null); }}
              title="Role Management"
              aria-label="Role Management"
              className={cn(
                'w-full flex items-center rounded-xl text-sm font-medium transition-all',
                sidebarCollapsed ? 'justify-center px-2 py-2.5' : 'gap-3 px-3 py-2.5',
                activeTab === 'admin' ? 'bg-indigo-50 text-indigo-600' : 'text-zinc-500 hover:bg-zinc-50'
              )}
            >
              <Users className="w-4 h-4 shrink-0" />
              {!sidebarCollapsed && <span className="truncate">Role Management</span>}
            </button>
          )}
        </nav>

        <div className={cn('border-t border-zinc-100 shrink-0', sidebarCollapsed ? 'p-2' : 'p-3')}>
          <div
            className={cn(
              'flex items-center mb-2',
              sidebarCollapsed ? 'flex-col gap-2 justify-center' : 'gap-3 px-1 py-2'
            )}
          >
            <div className="w-8 h-8 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center font-bold text-xs shrink-0">
              {user.username[0].toUpperCase()}
            </div>
            {!sidebarCollapsed && (
              <div className="flex-1 min-w-0">
                <p className="text-sm font-bold text-zinc-900 truncate">{user.username}</p>
                <p className="text-[10px] text-zinc-400 uppercase font-bold truncate">
                  {selectedEntity} • {(user.roles || []).join(', ')}
                </p>
              </div>
            )}
          </div>
          <button
            onClick={handleLogout}
            title="Log out"
            aria-label="Log out"
            className={cn(
              'w-full flex items-center rounded-xl text-sm font-medium text-red-500 hover:bg-red-50 transition-all',
              sidebarCollapsed ? 'justify-center px-2 py-2' : 'gap-3 px-3 py-2'
            )}
          >
            <LogOut className="w-4 h-4 shrink-0" />
            {!sidebarCollapsed && <span>Logout</span>}
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 min-w-0 overflow-y-auto">
        <header className="h-14 sm:h-16 bg-white border-b border-zinc-200 flex items-center justify-between px-4 lg:px-6 sticky top-0 z-10">
          <h2 className="text-sm font-bold text-zinc-900 uppercase tracking-wider">
            {activeTab === 'dashboard' ? 'Overview' : 
             activeTab === 'create' ? 'Template Designer' : 
             activeTab === 'templates' ? 'Workflow Templates' :
             activeTab === 'submit' ? 'Submit New Request' :
             activeTab === 'requests' ? 'Approvals & Requests' : 
             activeTab === 'procurement' ? 'Procurement Center' :
             activeTab === 'cost-centers' ? 'Cost centers' : 'Administration'}
          </h2>
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
              <Building2 className="w-4 h-4 text-zinc-400" />
              {hasMultipleEntities ? (
                <select
                  value={selectedEntity || ''}
                  onChange={(e) => handleEntityChange(e.target.value || null)}
                  className="h-8 min-w-[180px] rounded-lg border border-zinc-300 bg-white px-2.5 text-xs font-semibold text-zinc-700 outline-none focus:ring-2 focus:ring-indigo-500"
                >
                  <option value="" disabled>Select Active Entity</option>
                  {(user?.entities || []).map((ent) => (
                    <option key={ent} value={ent}>{ent}</option>
                  ))}
                </select>
              ) : (
                <span className="text-xs font-semibold text-zinc-600">{selectedEntity || 'No Entity'}</span>
              )}
            </div>
            <div className="h-8 w-px bg-zinc-200" />
            <span className="text-xs text-zinc-400 font-medium">{formatDateMYT(new Date())}</span>
          </div>
        </header>

        <div
          className={cn(
            "mx-auto w-full min-h-0",
            activeTab === 'procurement' || activeTab === 'cost-centers' || activeTab === 'requests' || activeTab === 'admin'
              ? "px-4 sm:px-6 lg:px-8 py-4 max-w-none h-[calc(100vh-4rem)] flex flex-col"
              : activeTab === 'submit' && selectedTemplateForRequest
              ? "flex flex-col h-[calc(100vh-4rem)] max-w-none px-4 sm:px-6 lg:px-8 py-4"
              : "px-6 py-5 max-w-none"
          )}
        >
          <AnimatePresence mode="wait">
            {activeTab === 'dashboard' && canSeeDashboard && !selectedTemplateForRequest && (
              <motion.div
                key="dashboard"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                <Dashboard 
                  requests={requests}
                  workflows={workflows}
                  user={user}
                  onStartRequest={(w) => {
                    setSelectedTemplateForRequest(w);
                    setActiveTab('submit');
                  }}
                  onViewRequest={(r) => {
                    setPreSelectedRequestId(r.id);
                    openRequestsTab();
                  }}
                  onViewAllRequests={openRequestsTab}
                />
              </motion.div>
            )}

            {activeTab === 'create' && canSeeTemplateDesigner && (
              <motion.div
                key="create"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                <WorkflowCreator 
                  onSuccess={() => { setActiveTab('templates'); fetchWorkflows(); }} 
                  availableRoles={roles}
                />
              </motion.div>
            )}
            
            {activeTab === 'templates' && canSeeWorkflowTemplates && !selectedTemplateForRequest && (
              <motion.div
                key="templates"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                <WorkflowList 
                  workflows={workflows}
                  user={user}
                  roles={roles}
                  onRefresh={fetchWorkflows}
                  onStartRequest={(w) => {
                    setSelectedTemplateForRequest(w);
                    setActiveTab('submit');
                  }}
                />
              </motion.div>
            )}

            {activeTab === 'submit' && (
              <motion.div
                key="submit"
                className={selectedTemplateForRequest ? "flex flex-col flex-1 min-h-0 h-full" : undefined}
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                {!selectedTemplateForRequest ? (
                  <TemplateSelector 
                    templates={workflows} 
                    onSelect={(w) => setSelectedTemplateForRequest(w)} 
                  />
                ) : (
                  <div className="flex flex-col flex-1 min-h-0 h-full gap-3">
                    <button 
                      type="button"
                      onClick={() => setSelectedTemplateForRequest(null)}
                      className="shrink-0 text-sm text-indigo-600 font-semibold flex items-center gap-1 hover:underline w-fit"
                    >
                      <ChevronRight className="w-4 h-4 rotate-180" />
                      Change Template
                    </button>
                    <div className="flex-1 min-h-0 flex flex-col">
                      <WorkflowRequestCreator 
                        template={selectedTemplateForRequest} 
                        entity={selectedEntity || undefined}
                        onSuccess={() => { openRequestsTab(); }} 
                      />
                    </div>
                  </div>
                )}
              </motion.div>
            )}

            {activeTab === 'requests' && (
              <motion.div
                key="requests"
                className="flex flex-col flex-1 min-h-0"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                <div className="mb-3 text-xs text-zinc-500 shrink-0">
                  Approvals view shows requests across your accessible entities. Submission always uses the active entity in the header.
                </div>
                <div className="flex-1 min-h-0 flex flex-col">
                  <WorkflowRequestList 
                    requests={requests} 
                    user={user}
                    onRefresh={fetchRequests} 
                    preSelectedRequestId={preSelectedRequestId}
                    onClearPreSelected={() => setPreSelectedRequestId(null)}
                  />
                </div>
              </motion.div>
            )}

            {activeTab === 'procurement' && hasPermission('view_procurement_center') && (
              <motion.div
                key="procurement"
                className="flex flex-col flex-1 min-h-0"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                <ProcurementCenter user={user} />
              </motion.div>
            )}

            {activeTab === 'cost-centers' && hasPermission('manage_cost_centers') && (
              <motion.div
                key="cost-centers"
                className="flex flex-col flex-1 min-h-0"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                <CostCentersAdminPage entity={selectedEntity} />
              </motion.div>
            )}

            {activeTab === 'admin' && user.roles?.some(r => r.toLowerCase() === 'admin') && (
              <motion.div
                key="admin"
                className="flex flex-col flex-1 min-h-0"
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -10 }}
              >
                <RoleManager roles={roles} onRefresh={fetchRoles} />
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </main>
    </div>
  );
}
