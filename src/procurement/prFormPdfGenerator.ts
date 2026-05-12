/**
 * F-PU-003 PR PDF (jsPDF) � extracted from App for reuse by maintenance scripts.
 * Keep layout aligned with the in-app PR form.
 */
import { jsPDF } from "jspdf";
import type { RequestApproval, WorkflowRequest, WorkflowStep } from "../types";

function normalizeWorkflowRequestStatus(s: string | undefined) {
  return (s ?? "").toString().trim().toLowerCase();
}

const ENTITY_LEGAL_NAMES: Record<string, string> = {
  GCBCM: "GCB COCOA MALAYSIA SDN BHD",
  GCCM: "GUAN CHONG COCOA MANUFACTURER SDN BHD",
  CCI: "GCB COCOA COTE D'IVOIRE",
  GCBCS: "GCB COCOA SINGAPORE PTE LTD",
  ACI: "PT. ACI COCOA INDONESIA",
};

function entityLegalDisplayName(code: string | undefined | null): string {
  if (code == null || !String(code).trim()) return "-";
  const key = String(code).trim().toUpperCase();
  return ENTITY_LEGAL_NAMES[key] || String(code).trim();
}

function procurementTaxLabelForEntity(code: string | undefined | null): string {
  const k = String(code ?? "").trim().toUpperCase();
  if (k === "GCCM" || k === "GCBCM") return "SST";
  if (k === "CCI") return "TVA";
  return "Tax";
}

function clampUnitRate(n: number): number {
  if (Number.isNaN(n) || n < 0) return 0;
  if (n > 1) return 1;
  return n;
}

function roundMoney2(n: number): number {
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 100) / 100;
}

function lineItemQtyForDisplay(item: any): number {
  const max = parseFloat(String(item?.["Max quantity"] ?? item?.["Max Quantity"] ?? ""));
  if (!Number.isNaN(max) && max > 0) return max;
  const q = parseFloat(String(item?.["Quantity"] ?? item?.quantity ?? 0));
  if (!Number.isNaN(q) && q > 0) return q;
  const min = parseFloat(String(item?.["Min quantity"] ?? item?.["Min Quantity"] ?? ""));
  if (!Number.isNaN(min)) return min;
  return 0;
}

function procurementMoneyTotals(req: Pick<WorkflowRequest, "line_items" | "tax_rate" | "discount_rate">) {
  let subtotal = 0;
  for (const item of req.line_items || []) {
    const qty = lineItemQtyForDisplay(item);
    const price = parseFloat(String(item?.["Unit Price"] ?? item?.["Price"] ?? item?.["Amount"] ?? 0));
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

const REMARKS_LINE_COL = "Remarks";
const LINE_ITEM_REMARKS_KEY = "_lineRemarks";

const lineItemRemarksDisplay = (item: any) =>
  String(item?.[REMARKS_LINE_COL] ?? item?.["Remarks (Purpose)"] ?? "").trim();

function prSuggestedSupplierDisplay(request: Pick<WorkflowRequest, "suggested_supplier" | "line_items">): string {
  const docLevel = (request.suggested_supplier ?? "").toString().trim();
  if (docLevel) return docLevel;
  for (const item of request.line_items || []) {
    const v = String(
      item?.["Suggested Supplier"] ?? item?.["suggested supplier"] ?? item?.["Supplier"] ?? ""
    ).trim();
    if (v) return v;
  }
  return "";
}

function toUpperSerial(value: string | number | null | undefined): string {
  return String(value ?? "").toUpperCase();
}

function displayRequestSerial(request: Pick<WorkflowRequest, "formatted_id" | "id">): string {
  return toUpperSerial(request.formatted_id || request.id);
}

function isSignatureImageDataUrl(s: string | undefined | null): boolean {
  if (!s || typeof s !== "string") return false;
  const t = s.trim();
  if (!/^data:image\/(png|jpeg|jpg);base64,/i.test(t)) return false;
  return t.length >= 80;
}

function formatSignatureProofTimestamp(iso: string | undefined | null): string {
  if (!iso) return "";
  try {
    return new Date(iso).toLocaleString();
  } catch {
    return "";
  }
}

const stepApproverRoleLowerAt = (steps: WorkflowStep[] | undefined, stepIndex: number): string =>
  (steps?.[stepIndex]?.approverRole ?? "").toString().trim().toLowerCase();

const approvalStepRoleLower = (a: RequestApproval, steps: WorkflowStep[] | undefined): string => {
  const snap = (a.approver_role_snapshot ?? "").toString().trim().toLowerCase();
  if (snap) return snap;
  return stepApproverRoleLowerAt(steps, a.step_index);
};

const workflowApprovedApprovals = (request: WorkflowRequest): RequestApproval[] => {
  const list = (request.approvals || []).filter((a) => normalizeWorkflowRequestStatus(a.status) === "approved");
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

function prHodApprovalForPdf(request: WorkflowRequest): RequestApproval | undefined {
  const steps = request.template_steps || [];
  const list = workflowApprovedApprovals(request);
  const hod = [...list].reverse().find((a) => approvalStepRoleLower(a, steps) === "approver");
  if (hod) return hod;
  const nonChecker = list.filter((a) => approvalStepRoleLower(a, steps) !== "checker");
  if (nonChecker.length) return nonChecker[nonChecker.length - 1];
  return undefined;
}

function pdfDesignationLines(
  doc: jsPDF,
  prefix: string,
  designation: string | null | undefined,
  maxW: number,
  fontSize = 7
) {
  const d = (designation ?? "").toString().trim();
  const s = d ? `${prefix}${d}` : `${prefix}\u2014`;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(fontSize);
  return doc.splitTextToSize(s, maxW);
}

/** F-PU-003 style form PDF — only for procurement Purchase Request (see isPR_Only). Landscape for wider table. */
export function buildProcurementPrFormPdfDoc(request: WorkflowRequest) {
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageW = 297;
  const pageH = 210;
  const left = 10;
  const right = pageW - 10;
  const tableLeft = left;
  const tableRight = right;
  const pageBottomSafe = pageH - 12;
  /** Column widths (mm); sum === tableRight − tableLeft (277mm). Supplier is in header, not grid. */
  const colWidths = [10, 108, 16, 36, 36, 71];
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
  const totalsBlockH = 26;
  /** Below chrome (incl. suggested supplier line). */
  const tableTopY = 42;
  const taxLabel = procurementTaxLabelForEntity(request.entity);
  const PR_MAX_ROWS_PER_PAGE = 5;

  /** Footer signature box is always drawn on every page. */
  const signatureBlockH = 56;
  const signatureBoxInnerH = 52;
  const signTopFixed = pageBottomSafe - signatureBoxInnerH;
  const tableBottomSafe = signTopFixed - 4;

  const getItemValue = (item: any, keys: string[]) => {
    const foundKey = Object.keys(item || {}).find((k) => keys.includes(k.toLowerCase()));
    if (!foundKey) return '';
    return String(item[foundKey] ?? '');
  };

  const lineItems = request.line_items || [];
  const money = procurementMoneyTotals(request);
  const { subtotal, discountRate, discountAmount, taxRate, taxAmount: tax, total } = money;

  const headerLabels = [
    'No',
    'Item',
    'Qty',
    'Cost Center Account No.',
    'Request to be delivered on',
    'Remarks'
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

  type RowBlock = { kind: 'data'; cells: string[][]; h: number } | { kind: 'totals'; h: number };

  const dataBlocks: RowBlock[] = [];

  lineItems.forEach((item, idx) => {
    doc.setFontSize(bodyFont);
    const qty = getItemValue(item, ['quantity', 'qty']);
    const itemName = getItemValue(item, ['item', 'description']);
    const cc = getItemValue(item, ['cost center', 'cost center account no.', 'cost centre']) || request.cost_center || '-';
    const eta = getItemValue(item, ['request to be delivered on', 'delivery date', 'date of delivery']);
    /** PR grid "Remarks" / "Remarks (Purpose)" only — not `_lineRemarks` line notes. */
    const tableRemarks = lineItemRemarksDisplay(item);
    const lineOnlyRemarks = item && item[LINE_ITEM_REMARKS_KEY] ? String(item[LINE_ITEM_REMARKS_KEY]).trim() : '';
    const itemCellText = lineOnlyRemarks ? `${itemName}\n\n${lineOnlyRemarks}` : itemName;

    const cells: string[][] = [
      wrap(String(idx + 1), 0, bodyFont),
      wrap(itemCellText, 1, bodyFont),
      wrap(qty || '-', 2, bodyFont),
      wrap(cc, 3, bodyFont),
      wrap(eta, 4, bodyFont),
      wrap(tableRemarks || '-', 5, bodyFont),
    ];
    const linesPerCol = cells.map((lines) => lines.length);
    const h = Math.max(minRowH, Math.max(...linesPerCol, 1) * lineStepBody + padY * 2);
    dataBlocks.push({ kind: 'data', cells, h });
  });

  const drawPageChrome = (tableTop: number) => {
    doc.setLineWidth(0.3);
    doc.rect(left, 8, right - left, pageH - 16);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.text(entityLegalDisplayName(request.entity).toUpperCase(), pageW / 2, 15, { align: 'center' });
    doc.setFontSize(9);
    doc.text('PURCHASE REQUISITION FORM: F-PU-003 (REV 0)', pageW / 2, 23, { align: 'center' });
    doc.setFontSize(8);
    doc.text('Department :', 14, 28);
    doc.setFont('helvetica', 'normal');
    doc.text(request.department || '-', 34, 28);
    doc.text(`Date: ${new Date(request.created_at).toLocaleDateString()}`, tableRight - padX, 28, { align: 'right' });
    doc.setFont('helvetica', 'bold');
    doc.text('Suggested supplier:', 14, 33);
    doc.setFont('helvetica', 'normal');
    const supLines = doc.splitTextToSize(prSuggestedSupplierDisplay(request) || '—', tableRight - 52 - padX);
    let sy = 33;
    for (let i = 0; i < Math.min(2, supLines.length); i++) {
      doc.text(supLines[i], 50, sy);
      sy += 3.5;
    }
    doc.text(`PR No: ${displayRequestSerial(request)}`, tableRight - padX, 38, { align: 'right' });
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

  const drawSignatureBlock = (signTop: number) => {
    const signBottom = signTop + signatureBoxInnerH;
    const hodApproval = prHodApprovalForPdf(request);

    doc.rect(left, signTop, right - left, signBottom - signTop);
    const mid = (left + right) / 2;
    doc.line(mid, signTop, mid, signBottom);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text('Requested By:', left + 2, signTop + 6);
    doc.text('Approved By:', mid + 2, signTop + 6);

    const sigPadX = 3;
    const sigMaxW = mid - left - sigPadX * 2;
    const sigMaxH = 22;
    const sigImgY = signTop + 8;

    const reqSigUrl = isSignatureImageDataUrl(request.requester_signature) ? request.requester_signature! : '';
    if (reqSigUrl) {
      addSignatureImageFit(reqSigUrl, left + sigPadX, sigImgY, sigMaxW, sigMaxH);
    } else if (request.requester_signed_at) {
      doc.setFontSize(7);
      doc.text('Electronic signature', left + sigPadX, sigImgY + 5);
      doc.text(formatSignatureProofTimestamp(request.requester_signed_at), left + sigPadX, sigImgY + 11);
      doc.setFontSize(8);
    }

    const apprSigUrl =
      hodApproval?.approver_signature && isSignatureImageDataUrl(hodApproval.approver_signature)
        ? hodApproval.approver_signature!
        : '';
    if (apprSigUrl) {
      addSignatureImageFit(apprSigUrl, mid + sigPadX, sigImgY, sigMaxW, sigMaxH);
    } else if (hodApproval && hodApproval.status.toLowerCase() === 'approved') {
      doc.setFontSize(7);
      doc.text('Electronic signature', mid + sigPadX, sigImgY + 5);
      doc.text(formatSignatureProofTimestamp(hodApproval.created_at), mid + sigPadX, sigImgY + 11);
      doc.setFontSize(8);
    }

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.text(`Name : ${request.requester_name || ''}`, left + 2, signBottom - 15);
    let yDesPr = signBottom - 11;
    pdfDesignationLines(doc, 'Designation: ', request.requester_designation, sigMaxW - 2, 6.5).forEach((ln) => {
      doc.text(ln, left + 2, yDesPr);
      yDesPr += 3.2;
    });
    doc.setFontSize(8);
    doc.text(`Name: ${hodApproval?.approver_name || ''}`, mid + 2, signBottom - 15);
    yDesPr = signBottom - 11;
    pdfDesignationLines(doc, 'Designation: ', hodApproval?.approver_designation ?? null, sigMaxW - 2, 6.5).forEach((ln) => {
      doc.text(ln, mid + 2, yDesPr);
      yDesPr += 3.2;
    });
    const proxyName = String(hodApproval?.signed_by_name ?? '').trim();
    if (proxyName) {
      doc.setFontSize(7);
      doc.setTextColor(67, 56, 202);
      doc.text(`Signed by proxy: ${proxyName}`, mid + 2, yDesPr);
      doc.setTextColor(0, 0, 0);
      doc.setFontSize(8);
      yDesPr += 3.2;
    }
    doc.setFontSize(8);
    doc.text(new Date(request.created_at).toLocaleDateString(), left + 2, signBottom - 2);
    doc.text(
      hodApproval?.created_at
        ? new Date(hodApproval.created_at).toLocaleDateString()
        : new Date().toLocaleDateString(),
      mid + 2,
      signBottom - 2
    );
  };

  const drawTotalsBlock = (yTop: number) => {
    const yBot = yTop + totalsBlockH;
    doc.line(tableLeft, yBot, tableRight, yBot);
    const splitAt = colWidths.length - 2;
    const lastColStart = colXs[colWidths.length - 1];
    doc.line(colXs[splitAt], yTop, colXs[splitAt], yBot);
    doc.line(lastColStart, yTop + 20, tableRight, yTop + 20);
    doc.setFontSize(8);
    doc.text('Subtotal', colXs[splitAt] + padX, yTop + 5);
    doc.text(subtotal.toFixed(2), tableRight - padX, yTop + 5, { align: 'right' });
    doc.text(`Discount (${(discountRate * 100).toFixed(0)}%)`, colXs[splitAt] + padX, yTop + 11);
    doc.text(discountAmount.toFixed(2), tableRight - padX, yTop + 11, { align: 'right' });
    doc.text(`${taxLabel} ${(taxRate * 100).toFixed(0)}%`, colXs[splitAt] + padX, yTop + 17);
    doc.text(`${(tax || 0).toFixed(2)}`, tableRight - padX, yTop + 17, { align: 'right' });
    doc.text((request.currency?.trim() || '—').toUpperCase(), lastColStart + padX, yBot - 3);
    doc.text(total.toFixed(2), tableRight - padX, yBot - 3, { align: 'right' });
    return yBot;
  };

  const drawNewPage = () => {
    drawPageChrome(tableTopY);
    const headerBottom = drawHeaderRow(tableTopY);
    drawVerticalRules(tableTopY, headerBottom);
    return headerBottom;
  };

  let pageIndex = 0;
  let rowOnPage = 0;
  let y = tableTopY;
  y = drawNewPage();

  for (let i = 0; i < dataBlocks.length; i++) {
    const block = dataBlocks[i];
    if (block.kind !== 'data') continue;

    const needNewPage =
      rowOnPage >= PR_MAX_ROWS_PER_PAGE || (y + block.h > tableBottomSafe);
    if (needNewPage) {
      // Close previous page with signature before continuing.
      drawSignatureBlock(signTopFixed);
      doc.addPage('a4', 'l');
      pageIndex += 1;
      rowOnPage = 0;
      y = drawNewPage();
    }

    y = drawDataRow(y, block.h, block.cells);
    drawVerticalRules(tableTopY, y);
    rowOnPage += 1;
  }

  // Totals: place on the last page above signatures; if it doesn't fit, move totals to a new page.
  const totalsNeedNewPage = y + totalsBlockH > tableBottomSafe;
  if (totalsNeedNewPage) {
    drawSignatureBlock(signTopFixed);
    doc.addPage('a4', 'l');
    pageIndex += 1;
    y = drawNewPage();
  }
  y = drawTotalsBlock(y);
  drawVerticalRules(tableTopY, y);
  drawSignatureBlock(signTopFixed);

  return doc;
};

