/**
 * utils/excelProcessor.ts
 * 100% client-side Excel processing.
 *
 * Read:  xlsx (SheetJS)
 * Write: exceljs (styled .xlsx)
 * Save:  file-saver
 *
 * Output: ONE Excel file with 2 sheets:
 *  - Sheet 1: สรุปผล (แบบรูป)
 *  - Sheet 2: ข้อมูลทั้งหมด
 */

import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// ───────────────────────────────────────────────────────────────────────────────
// Overdue buckets
// ───────────────────────────────────────────────────────────────────────────────

export const OVERDUE_BUCKETS = [
  'ไม่เกิน 30 วัน',
  '31-60 วัน',
  '61-90 วัน',
  '91-120 วัน',
  '121-150 วัน',
  '151 วันขึ้นไป',
] as const;

export type OverdueBucket = typeof OVERDUE_BUCKETS[number];

// ───────────────────────────────────────────────────────────────────────────────
// Types
// ───────────────────────────────────────────────────────────────────────────────

export interface CollectionRecord {
  index: number | string;
  employeeId: string;
  branch: string;

  saleDate: string; // dd/mm/yyyy
  overdueDays: number | string;
  overdueBucket: OverdueBucket;

  contactDate: string; // dd/mm/yyyy
  contractNo: string;
  fullName: string;

  totalPaidAmount: number; // from export "ยอดรับสุทธิ"
  status: 'SELF_PAID' | 'CALL_PAID' | 'UNPAID';
}

export interface BucketStat {
  bucket: OverdueBucket;
  total: number; // งานที่ได้รับทั้งหมด
  callPaid: number; // ทำได้ = CALL_PAID
  selfPaid: number; // มาเอง = SELF_PAID
  unpaid: number; // ไม่มาจ่าย = UNPAID
}

export interface EmployeeSummary {
  employeeId: string;
  buckets: BucketStat[];
  total: number;
  callPaid: number;
  selfPaid: number;
  unpaid: number;
}

export interface ProcessingResult {
  reportTitle: string;
  summary: {
    total: number;
    selfPaid: number;
    callPaid: number;
    unpaid: number;
    successRate: number; // ((total - selfPaid) / callPaid) * 100  (ตามสูตรที่คุณให้)
  };
  records: CollectionRecord[];
  employeeSummary: EmployeeSummary[];
}

// ───────────────────────────────────────────────────────────────────────────────
// Column aliases (Thai headers vary by file)
// ───────────────────────────────────────────────────────────────────────────────

const ALIAS = {
  // Collection file
  INDEX: ['ลำดับ', 'ลำดบ'],
  EMPLOYEE_ID: ['รหัสพนักงาน', 'พนักงาน', 'ชื่อพนักงาน', 'รหัส พนง.', 'รหัสพนง'],
  BRANCH: ['สาขา', 'สาขา/หน่วย'],
  SALE_DATE: ['วันที่ขาย', 'วันขาย'],
  OVERDUE_DAYS: ['วันขาดการติดต่อ', 'ขาดการติดต่อ', 'วันขาด'],
  CONTACT_DATE: ['วันที่ติดต่อ', 'วันติดต่อ'],
  CONTRACT_NO: ['เลขที่สัญญา', 'เลขทีสัญญา', 'เลขสัญญา'],
  FULL_NAME: ['ชื่อ-นามสกุล', 'ชื่อ - สกุล', 'ชื่อสกุล', 'ชื่อ-สกุล'],
  // ผลการตามอาจอยู่ใต้คอลัมน์ "ผลตาม/ผลการตาม/ลูกค้ามาจ่าย" แล้วค่ามีคำว่า "ลูกค้ามาจ่าย"
  RESULT_TEXT: ['ผลตาม', 'ผลการตาม', 'ผลการติดตาม', 'ลูกค้ามาจ่าย'],

  // Export file
  EXPORT_CONTRACT_NO: ['เลขที่สัญญา', 'เลขทีสัญญา', 'เลขสัญญา'],
  NET_RECEIVED: ['ยอดรับสุทธิ', 'ยอดรับสุทธิ(บาท)', 'ยอดรับสุทธิ (บาท)', 'net_received', 'netReceived'],
  EXPORT_EMPLOYEE_ID: ['พนง.เก็บเงิน'],
};

// ───────────────────────────────────────────────────────────────────────────────
// Helpers
// ───────────────────────────────────────────────────────────────────────────────

/**
 * IMPORTANT: contract numbers often have leading zeros.
 * Excel / SheetJS may parse them as number => leading zeros are LOST.
 *
 * Fix:
 * - read with raw: true
 * - keep a display value for export/UI
 * - build a normalized key for matching
 * - also keep fallback key without leading zeros in exportMap (for matching)
 */
const CONTRACT_LEN = 14;

const cleanContractDisplay = (value: unknown): string => {
  if (value === null || value === undefined) return '';

  let s = String(value).trim().replace(/\u200b/g, '');
  s = s.replace(/\s+/g, '');
  s = s.replace(/\.0+$/, '');
  return s;
};

const normalizeContractKey = (value: unknown): string => {
  let s = cleanContractDisplay(value).toUpperCase();
  if (!s) return '';

  s = s.replace(/[^A-Z0-9]/g, '');
  if (!s) return '';

  if (/^\d+$/.test(s) && s.length < CONTRACT_LEN) {
    s = s.padStart(CONTRACT_LEN, '0');
  }

  return s;
};

const contractKeyNoZero = (contract: string) => {
  if (!/^\d+$/.test(contract)) return contract;
  return contract.replace(/^0+/, '');
};

const parseAmount = (value: unknown): number => {
  if (value === null || value === undefined) return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  let s = String(value).trim();
  if (!s || s === '-') return 0;

  // (1,234.00) => -1234.00
  const negative = /^\(.*\)$/.test(s);
  if (negative) s = s.slice(1, -1);

  // remove currency/comma/spaces
  s = s.replace(/[฿,\s]/g, '');

  const n = Number(s);
  if (!Number.isFinite(n)) return 0;
  return negative ? -n : n;
};

const formatDate = (value: unknown): string => {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') {
    const d = XLSX.SSF.parse_date_code(value);
    if (d) {
      const dd = String(d.d).padStart(2, '0');
      const mm = String(d.m).padStart(2, '0');
      return `${dd}/${mm}/${d.y}`;
    }
  }
  const s = String(value).trim();
  return s.split(' ')[0]; // "15/01/2026 00:00:00" -> "15/01/2026"
};

export const getOverdueBucket = (raw: unknown): OverdueBucket => {
  const n = typeof raw === 'number' ? raw : parseFloat(String(raw ?? '').replace(/,/g, ''));
  if (!Number.isFinite(n)) return 'ไม่เกิน 30 วัน';
  if (n <= 30) return 'ไม่เกิน 30 วัน';
  if (n <= 60) return '31-60 วัน';
  if (n <= 90) return '61-90 วัน';
  if (n <= 120) return '91-120 วัน';
  if (n <= 150) return '121-150 วัน';
  return '151 วันขึ้นไป';
};

const pctNum = (num: number, denom: number): number => (denom === 0 ? 0 : num / denom);

const firstNonEmpty = (...vals: unknown[]): string => {
  for (const v of vals) {
    const s = String(v ?? '').trim();
    if (s !== '' && s.toLowerCase() !== 'nan' && s.toLowerCase() !== 'undefined') return s;
  }
  return '';
};

const getByAliases = (row: Record<string, unknown>, keys: string[]) => {
  for (const k of keys) {
    if (k in row) return row[k];
  }
  return undefined;
};

const cleanTitleCell = (value: unknown): string =>
  String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();

const extractEmployeeNameFromTitle = (value: unknown): string => {
  const text = cleanTitleCell(value);
  if (!text) return '';

  const collectionMatch = text.match(/รายงานผลการตาม\s+(.+?)\s+เดือน/i);
  if (collectionMatch?.[1]) return collectionMatch[1].trim();

  const exportMatch = text.match(/^(.*?)\s+สาขา\s*\d+/i);
  if (exportMatch?.[1]) return exportMatch[1].trim();

  return '';
};

/**
 * Scan first 25 rows for any header value; return 0-based row index.
 */
const findHeaderRow = (ws: XLSX.WorkSheet, colName: string): number => {
  const ref = ws['!ref'] || 'A1:A1';
  const range = XLSX.utils.decode_range(ref);
  for (let R = range.s.r; R <= Math.min(range.e.r, 25); R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell?.v && String(cell.v).trim() === colName) return R;
    }
  }
  return 0;
};

const findHeaderRowAny = (ws: XLSX.WorkSheet, colNames: string[]): number => {
  for (const name of colNames) {
    const r = findHeaderRow(ws, name);
    if (r !== 0) return r;
  }
  return 0;
};

// ───────────────────────────────────────────────────────────────────────────────
// Build Employee Summary (for report)
// ───────────────────────────────────────────────────────────────────────────────

const buildEmployeeSummary = (records: CollectionRecord[]): EmployeeSummary[] => {
  const empMap = new Map<string, Map<OverdueBucket, BucketStat>>();

  for (const rec of records) {
    const eid = rec.employeeId || '-';
    if (!empMap.has(eid)) empMap.set(eid, new Map());
    const bmap = empMap.get(eid)!;

    if (!bmap.has(rec.overdueBucket)) {
      bmap.set(rec.overdueBucket, { bucket: rec.overdueBucket, total: 0, callPaid: 0, selfPaid: 0, unpaid: 0 });
    }

    const stat = bmap.get(rec.overdueBucket)!;
    stat.total++;
    if (rec.status === 'CALL_PAID') stat.callPaid++;
    else if (rec.status === 'SELF_PAID') stat.selfPaid++;
    else stat.unpaid++;
  }

  const out: EmployeeSummary[] = [];
  for (const [employeeId, bmap] of empMap.entries()) {
    const buckets = OVERDUE_BUCKETS.map(
      (b) => bmap.get(b) ?? { bucket: b, total: 0, callPaid: 0, selfPaid: 0, unpaid: 0 }
    );

    const agg = buckets.reduce(
      (acc, s) => ({
        total: acc.total + s.total,
        callPaid: acc.callPaid + s.callPaid,
        selfPaid: acc.selfPaid + s.selfPaid,
        unpaid: acc.unpaid + s.unpaid,
      }),
      { total: 0, callPaid: 0, selfPaid: 0, unpaid: 0 }
    );

    out.push({ employeeId, buckets, ...agg });
  }

  return out.sort((a, b) => a.employeeId.localeCompare(b.employeeId, 'th'));
};

// ───────────────────────────────────────────────────────────────────────────────
// Main Processor
// ───────────────────────────────────────────────────────────────────────────────

export const processExcelData = (collectionBuffer: ArrayBuffer, exportBuffer: ArrayBuffer): ProcessingResult => {
  // 1) Parse Collection
  const wb1 = XLSX.read(collectionBuffer, { type: 'array' });
  const ws1 = wb1.Sheets[wb1.SheetNames[0]];
  const collectionTitle = cleanTitleCell(ws1['A1']?.v);
  const collectionEmployeeName = extractEmployeeNameFromTitle(collectionTitle);
  const hdr1 = findHeaderRowAny(ws1, ALIAS.CONTRACT_NO);

  const collectionRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws1, {
    range: hdr1,
    defval: '',
    raw: true, // IMPORTANT
  });

  // 2) Parse Export
  const wb2 = XLSX.read(exportBuffer, { type: 'array' });
  const ws2 = wb2.Sheets[wb2.SheetNames[0]];
  const exportSubtitle = cleanTitleCell(ws2['A2']?.v);
  const exportEmployeeName = extractEmployeeNameFromTitle(exportSubtitle);
  const hdr2 = findHeaderRowAny(ws2, ALIAS.EXPORT_CONTRACT_NO);

  const exportRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws2, {
    range: hdr2,
    defval: '',
    raw: true, // IMPORTANT
  });

  const exportEmployeeId = firstNonEmpty(
    ...exportRows.map((row) => getByAliases(row, ALIAS.EXPORT_EMPLOYEE_ID))
  );

  // 3) ExportMap: contractNo -> totalPaidAmount (use "ยอดรับสุทธิ" only)
  const exportMap = new Map<string, number>();

  const addExportAmount = (contract: string, amount: number) => {
    if (!contract) return;
    exportMap.set(contract, (exportMap.get(contract) ?? 0) + amount);

    // fallback key without leading zeros
    const noZero = contractKeyNoZero(contract);
    if (noZero) exportMap.set(noZero, (exportMap.get(noZero) ?? 0) + amount);
  };

  for (const row of exportRows) {
    const rawContract = getByAliases(row, ALIAS.EXPORT_CONTRACT_NO);
    const contract = normalizeContractKey(rawContract);
    if (!contract) continue;

    const netReceived = parseAmount(getByAliases(row, ALIAS.NET_RECEIVED));
    addExportAmount(contract, netReceived);
  }

  // 4) Group Collection by contract (1 contract => 1 record)
  const grouped = new Map<
    string,
    { displayContract: string; rows: Record<string, unknown>[] }
  >();
  for (const row of collectionRows) {
    const rawContract = getByAliases(row, ALIAS.CONTRACT_NO);
    const contractKey = normalizeContractKey(rawContract);
    if (!contractKey) continue;
    const displayContract = cleanContractDisplay(rawContract);

    if (!grouped.has(contractKey)) {
      grouped.set(contractKey, { displayContract, rows: [] });
    }

    const entry = grouped.get(contractKey)!;
    if (!entry.displayContract && displayContract) entry.displayContract = displayContract;
    entry.rows.push(row);
  }

  const records: CollectionRecord[] = [];
  const fallbackEmployeeName = collectionEmployeeName || exportEmployeeName;
  let selfPaid = 0,
    callPaid = 0,
    unpaid = 0;

  for (const [contractKey, entry] of grouped.entries()) {
    const { displayContract, rows } = entry;
    const pick = (keys: string[]) => firstNonEmpty(...rows.map((r) => getByAliases(r, keys)));

    const rawOverdue = pick(ALIAS.OVERDUE_DAYS);
    const overdueDays: number | string =
      rawOverdue === '' ? '' : (Number.isFinite(Number(rawOverdue)) ? Number(rawOverdue) : rawOverdue);

    // SELF_PAID detection: column might contain text "ลูกค้ามาจ่าย"
    const isSelfPaid = rows.some((r) => {
      const v = String(getByAliases(r, ALIAS.RESULT_TEXT) ?? '').trim();
      return v === 'ลูกค้ามาจ่าย' || v.includes('ลูกค้ามาจ่าย');
    });

    const lookupNoZero = contractKeyNoZero(contractKey);

    const totalPaidAmount =
      exportMap.get(contractKey) ?? (lookupNoZero ? exportMap.get(lookupNoZero) : undefined) ?? 0;

    const isInExport = exportMap.has(contractKey) || (lookupNoZero ? exportMap.has(lookupNoZero) : false);

    let status: CollectionRecord['status'];
    if (isSelfPaid) {
      status = 'SELF_PAID';
      selfPaid++;
    } else if (isInExport) {
      status = 'CALL_PAID';
      callPaid++;
    } else {
      status = 'UNPAID';
      unpaid++;
    }

    records.push({
      index: pick(ALIAS.INDEX),
      employeeId: exportEmployeeId || pick(ALIAS.EMPLOYEE_ID) || fallbackEmployeeName,
      branch: pick(ALIAS.BRANCH),
      saleDate: formatDate(pick(ALIAS.SALE_DATE)),
      overdueDays,
      overdueBucket: getOverdueBucket(overdueDays),
      contactDate: formatDate(pick(ALIAS.CONTACT_DATE)),
      contractNo: displayContract || contractKey,
      fullName: pick(ALIAS.FULL_NAME),
      totalPaidAmount,
      status,
    });
  }

  // sort by index if numeric
  records.sort((a, b) => {
    const an = Number(a.index);
    const bn = Number(b.index);
    if (Number.isFinite(an) && Number.isFinite(bn)) return an - bn;
    return String(a.index ?? '').localeCompare(String(b.index ?? ''), 'th');
  });

  const employeeSummary = buildEmployeeSummary(records);

  const total = records.length;
  // ✅ ตามสูตรที่คุณระบุ: (ทำได้ / (งานทั้งหมด - มาเอง)) * 100
  const successRate = callPaid === 0 ? 0 : (callPaid / (total - selfPaid)) * 100;

  return {
    reportTitle: collectionTitle || 'สรุปผลงานโทร',
    summary: { total, selfPaid, callPaid, unpaid, successRate },
    records,
    employeeSummary,
  };
};

// ───────────────────────────────────────────────────────────────────────────────
// Excel styling helpers
// ───────────────────────────────────────────────────────────────────────────────

type ArgbColor = string;

const fill = (argb: ArgbColor): ExcelJS.Fill => ({
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb },
});

const border = (): Partial<ExcelJS.Borders> => ({
  top: { style: 'thin', color: { argb: 'FFBDBDBD' } },
  bottom: { style: 'thin', color: { argb: 'FFBDBDBD' } },
  left: { style: 'thin', color: { argb: 'FFBDBDBD' } },
  right: { style: 'thin', color: { argb: 'FFBDBDBD' } },
});

const applyHeaderStyle = (cell: ExcelJS.Cell, fillArgb: ArgbColor, fontColor = 'FF000000') => {
  cell.fill = fill(fillArgb);
  cell.font = { bold: true, color: { argb: fontColor }, size: 10 };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cell.border = border();
};

// ───────────────────────────────────────────────────────────────────────────────
// Export ONE workbook with 2 sheets (Summary first, Detail second)
// ───────────────────────────────────────────────────────────────────────────────

export const generateOneExcelFile = async (result: ProcessingResult): Promise<void> => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Collection Dashboard';
  wb.created = new Date();

  const wsSummary = wb.addWorksheet('สรุปผล');
  buildSummarySheetLikeImage(wsSummary, result, false);

  const wsPerformance = wb.addWorksheet('ผลการติดตามความสำเสร็จ');
  buildSummarySheetLikeImage(wsPerformance, result, true);

  const wsDetail = wb.addWorksheet('ข้อมูลทั้งหมด');
  buildDetailedSheet(wsDetail, result);

  const buf = await wb.xlsx.writeBuffer();
  const safeTitle = (result.reportTitle || 'สรุปผลการตาม')
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim();
  saveAs(
    new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    `${safeTitle || 'สรุปผลการตาม'}.xlsx`
  );
};

// ───────────────────────────────────────────────────────────────────────────────
// Sheet builders
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Summary sheet layout styled like your image.
 */
const buildSummarySheetLikeImage = (
  ws: ExcelJS.Worksheet,
  result: ProcessingResult,
  performanceOnly = false
) => {
  const C = {
    titleBg: 'FFB7CDE8',
    headerBg: 'FFB7CDE8',
    headerFont: 'FF000000',
    orange: 'FFFFD580',
    green: 'FFC6EFCE',
    gray: 'FFD9D9D9',
    subtotalBg: 'FF92D050',
    grandBg: 'FFE7E6E6',
    successBg: 'FFDDEBF7',
  };

  const titleText = performanceOnly ? 'ผลการติดตามความสำเสร็จ' : result.reportTitle || 'ผลงานแรงรัดโทร';
  const title = ws.addRow([titleText]);
  ws.mergeCells(1, 1, 1, performanceOnly ? 7 : 10);
  title.height = 22;
  title.getCell(1).font = { bold: true, size: 14 };
  title.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
  title.getCell(1).fill = fill(C.titleBg);
  title.getCell(1).border = border();

  // Column widths
  ws.columns = [
    { width: 16 },
    { width: 18 },
    { width: 18 },
    { width: 10 },
    { width: 12 },
    { width: 10 },
    { width: 16 },
    { width: 12 },
    { width: 18 },
    { width: 12 },
  ];

  if (performanceOnly) {
    const secondaryHeader = ws.addRow([
      'รหัสพนักงาน',
      'ขาดการติดต่อ',
      'งานที่ตามเองทั้งหมด',
      'ทำได้',
      '%ทำได้',
      'ตามจ่ายไม่ได้',
      '%ตามจ่ายไม่ได้',
    ]);
    secondaryHeader.height = 24;
    for (let col = 1; col <= 7; col++) {
      applyHeaderStyle(secondaryHeader.getCell(col), C.headerBg, C.headerFont);
    }

    const secondaryPercentCols = new Set([5, 7]);
    const styleSecondaryRow = (r: ExcelJS.Row, isSubTotal = false) => {
      for (let col = 1; col <= 7; col++) {
        const cell = r.getCell(col);
        cell.border = border();
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

        if (isSubTotal) {
          cell.fill = fill(C.subtotalBg);
          cell.font = { ...(cell.font ?? {}), bold: true };
        } else {
          if (col === 4 || col === 5) cell.fill = fill(C.orange);
          if (col === 6 || col === 7) cell.fill = fill(C.gray);
        }

        if (secondaryPercentCols.has(col)) cell.numFmt = '0.00%';
      }
    };

    for (const emp of result.employeeSummary) {
      for (const b of emp.buckets) {
        if (b.total === 0) continue;
        const baseTotal = b.total - b.selfPaid;

        const row = ws.addRow([
          emp.employeeId,
          b.bucket,
          baseTotal,
          b.callPaid,
          pctNum(b.callPaid, baseTotal),
          b.unpaid,
          pctNum(b.unpaid, baseTotal),
        ]);
        styleSecondaryRow(row);
      }

      const empBaseTotal = emp.total - emp.selfPaid;
      const subtotalRow = ws.addRow([
        emp.employeeId,
        'รวม',
        empBaseTotal,
        emp.callPaid,
        pctNum(emp.callPaid, empBaseTotal),
        emp.unpaid,
        pctNum(emp.unpaid, empBaseTotal),
      ]);
      styleSecondaryRow(subtotalRow, true);
    }

    ws.views = [{ state: 'frozen', ySplit: 2 }];
    return;
  }

  // Header row
  const header = ws.addRow([
    'รหัสพนักงาน',
    'ขาดการติดต่อ',
    'งานที่ได้รับทั้งหมด',
    'ทำได้',
    '%ทำได้',
    'ไม่มาจ่าย',
    '%ลูกค้าไม่มาจ่าย',
    'ลูกค้ามาจ่ายเอง',
    '%ลูกค้ามาจ่ายเอง',
    '100.00%',
  ]);
  header.height = 26;
  header.eachCell((cell) => applyHeaderStyle(cell, C.headerBg, C.headerFont));

  const percentCols = new Set([5, 7, 9, 10]);

  const styleRow = (r: ExcelJS.Row, isSubTotal = false, isGrand = false) => {
    r.eachCell((cell, col) => {
      cell.border = border();
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

      if (isSubTotal) {
        cell.fill = fill(C.subtotalBg);
        cell.font = { ...(cell.font ?? {}), bold: true };
      } else if (isGrand) {
        cell.fill = fill(C.grandBg);
        cell.font = { ...(cell.font ?? {}), bold: true };
      } else {
        // Column colors like image
        if (col === 4 || col === 5) cell.fill = fill(C.orange);
        if (col === 6 || col === 7) cell.fill = fill(C.gray);
        if (col === 8 || col === 9) cell.fill = fill(C.green);
      }

      if (percentCols.has(col)) cell.numFmt = '0.00%';
    });
  };

  for (const emp of result.employeeSummary) {
    // bucket rows
    for (const b of emp.buckets) {
      if (b.total === 0) continue;

      const r = ws.addRow([
        emp.employeeId,
        b.bucket,
        b.total,
        b.callPaid,
        pctNum(b.callPaid, b.total),
        b.unpaid,
        pctNum(b.unpaid, b.total),
        b.selfPaid,
        pctNum(b.selfPaid, b.total),
        1,
      ]);
      styleRow(r);
    }

    // subtotal
    const sub = ws.addRow([
      emp.employeeId,
      'รวม',
      emp.total,
      emp.callPaid,
      pctNum(emp.callPaid, emp.total),
      emp.unpaid,
      pctNum(emp.unpaid, emp.total),
      emp.selfPaid,
      pctNum(emp.selfPaid, emp.total),
      1,
    ]);
    styleRow(sub, true, false);
  }

  ws.views = [{ state: 'frozen', ySplit: 2 }];
};

const buildDetailedSheet = (ws: ExcelJS.Worksheet, result: ProcessingResult) => {
  ws.columns = [
    { header: 'ลำดับ', key: 'index', width: 8 },
    { header: 'รหัสพนักงาน', key: 'employeeId', width: 14 },
    { header: 'สาขา', key: 'branch', width: 14 },
    { header: 'เลขที่สัญญา', key: 'contractNo', width: 18 },
    { header: 'ชื่อ-นามสกุล', key: 'fullName', width: 24 },
    { header: 'วันที่ขาย', key: 'saleDate', width: 14 },
    { header: 'วันขาดการติดต่อ', key: 'overdueDays', width: 16 },
    { header: 'ช่วงขาดการติดต่อ', key: 'overdueBucket', width: 18 },
    { header: 'วันที่ติดต่อ', key: 'contactDate', width: 14 },
    { header: 'ยอดชำระรวม', key: 'totalPaidAmount', width: 16 },
    { header: 'ผลการตาม', key: 'status', width: 16 },
  ];

  // Header style
  const hdr = ws.getRow(1);
  hdr.height = 28;
  hdr.eachCell((cell) => {
    cell.fill = fill('FF1F3864');
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = border();
  });

  const STATUS_LABEL: Record<CollectionRecord['status'], string> = {
    SELF_PAID: 'จ่ายเอง',
    CALL_PAID: 'โทรตามมาจ่าย',
    UNPAID: 'ไม่จ่าย',
  };

  const STATUS_FILL: Record<CollectionRecord['status'], ArgbColor> = {
    SELF_PAID: 'FFD4EDFF',
    CALL_PAID: 'FFC6EFCE',
    UNPAID: 'FFFFC7CE',
  };

  for (const rec of result.records) {
    const row = ws.addRow({
      index: rec.index,
      employeeId: rec.employeeId,
      branch: rec.branch,
      contractNo: rec.contractNo,
      fullName: rec.fullName,
      saleDate: rec.saleDate,
      overdueDays: rec.overdueDays,
      overdueBucket: rec.overdueBucket,
      contactDate: rec.contactDate,
      totalPaidAmount: rec.totalPaidAmount,
      status: STATUS_LABEL[rec.status],
    });

    row.height = 18;
    row.eachCell((cell, colNum) => {
      cell.border = border();
      cell.font = { size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: colNum === 10 ? 'right' : 'left', wrapText: true };
      if (colNum === 10) cell.numFmt = '#,##0.00';
    });

    // status cell color
    const statusCell = row.getCell(11);
    statusCell.fill = fill(STATUS_FILL[rec.status]);
    statusCell.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  ws.views = [{ state: 'frozen', ySplit: 1 }];
};
