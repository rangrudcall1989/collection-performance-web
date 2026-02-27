'use client';

import { useCallback, useMemo, useState } from 'react';
import {
  processExcelData,
  generateOneExcelFile,
  ProcessingResult,
  CollectionRecord,
  OVERDUE_BUCKETS,
} from '../utils/excelProcessor';

// ─── Types ─────────────────────────────────────────────────────────────────────

type SortKey = keyof CollectionRecord;
type SortDir = 'asc' | 'desc';
type StatusFilter = 'ALL' | 'SELF_PAID' | 'CALL_PAID' | 'UNPAID';

// ─── Status config ──────────────────────────────────────────────────────────────

const STATUS_CONFIG = {
  SELF_PAID: { label: 'จ่ายเอง',        badge: 'badge-blue',  dot: 'dot-blue'  },
  CALL_PAID: { label: 'โทรตามมาจ่าย',   badge: 'badge-green', dot: 'dot-green' },
  UNPAID:    { label: 'ไม่จ่าย',         badge: 'badge-red',   dot: 'dot-red'   },
} as const;

// ─── Sub-components ────────────────────────────────────────────────────────────

function UploadZone({
  label,
  fileName,
  accentClass,
  onChange,
}: {
  label: string;
  fileName: string;
  accentClass: string;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <div className={`upload-card ${accentClass}`}>
      <h3 className="upload-label">{label}</h3>
      <label className="upload-zone">
        <input type="file" accept=".xlsx,.xls" onChange={onChange} className="sr-only" />
        {fileName ? (
          <div className="upload-done">
            <span className="upload-check">✓</span>
            <span className="upload-filename">{fileName}</span>
          </div>
        ) : (
          <div className="upload-idle">
            <span className="upload-icon">📂</span>
            <span className="upload-hint">คลิกหรือลากไฟล์มาวาง</span>
            <span className="upload-ext">.xlsx, .xls</span>
          </div>
        )}
      </label>
    </div>
  );
}

function SummaryCard({ value, label, colorClass }: { value: number; label: string; colorClass: string }) {
  return (
    <div className={`summary-card ${colorClass}`}>
      <p className="summary-value">{value.toLocaleString()}</p>
      <p className="summary-label">{label}</p>
    </div>
  );
}

function StatusBadge({ status }: { status: CollectionRecord['status'] }) {
  const cfg = STATUS_CONFIG[status];
  return (
    <span className={`badge ${cfg.badge}`}>
      <span className={`dot ${cfg.dot}`} />
      {cfg.label}
    </span>
  );
}

function SortIcon({ active, dir }: { active: boolean; dir: SortDir }) {
  if (!active) return <span className="sort-icon inactive">↕</span>;
  return <span className="sort-icon active">{dir === 'asc' ? '↑' : '↓'}</span>;
}

// ─── Summary Table (HTML table จริง) ───────────────────────────────────────────

function SummaryReportTable({ result }: { result: ProcessingResult }) {
  const rows = useMemo(() => {
    const out: Array<{
      key: string;
      employeeId: string;
      bucket: string;
      total: number;
      done: number;
      pctDone: number;
      selfPaid: number;
      pctSelf: number;
      unpaid: number;
      pctUnpaid: number;
      isSubtotal: boolean;
      isGrand: boolean;
    }> = [];

    let gtTotal = 0, gtDone = 0, gtSelf = 0, gtUnpaid = 0;

    for (const emp of result.employeeSummary) {
      for (const b of emp.buckets) {
        if (b.total === 0) continue;

        const done = b.callPaid; // "ทำได้" ในรูป = โทรตามมาจ่าย
        out.push({
          key: `${emp.employeeId}-${b.bucket}`,
          employeeId: emp.employeeId,
          bucket: b.bucket,
          total: b.total,
          done,
          pctDone: b.total ? done / b.total : 0,
          selfPaid: b.selfPaid,
          pctSelf: b.total ? b.selfPaid / b.total : 0,
          unpaid: b.unpaid,
          pctUnpaid: b.total ? b.unpaid / b.total : 0,
          isSubtotal: false,
          isGrand: false,
        });
      }

      // subtotal
      const empDone = emp.callPaid;
      out.push({
        key: `${emp.employeeId}-subtotal`,
        employeeId: emp.employeeId,
        bucket: 'รวม',
        total: emp.total,
        done: empDone,
        pctDone: emp.total ? empDone / emp.total : 0,
        selfPaid: emp.selfPaid,
        pctSelf: emp.total ? emp.selfPaid / emp.total : 0,
        unpaid: emp.unpaid,
        pctUnpaid: emp.total ? emp.unpaid / emp.total : 0,
        isSubtotal: true,
        isGrand: false,
      });

      gtTotal += emp.total;
      gtDone += empDone;
      gtSelf += emp.selfPaid;
      gtUnpaid += emp.unpaid;
    }

    // grand total
    out.push({
      key: `grand-total`,
      employeeId: 'Grand Total',
      bucket: '',
      total: gtTotal,
      done: gtDone,
      pctDone: gtTotal ? gtDone / gtTotal : 0,
      selfPaid: gtSelf,
      pctSelf: gtTotal ? gtSelf / gtTotal : 0,
      unpaid: gtUnpaid,
      pctUnpaid: gtTotal ? gtUnpaid / gtTotal : 0,
      isSubtotal: false,
      isGrand: true,
    });

    return out;
  }, [result]);

  return (
    <div className="report-wrap">
      <div className="report-title">
        <div className="report-title-main">ผลงานแรงรัดโทร มกราคม 2569</div>
        <div className="report-title-sub">สรุปแยกตามรหัสพนักงาน และช่วงขาดการติดต่อ</div>
      </div>

      <div className="report-table-scroll">
        <table className="report-table">
          <thead>
            <tr>
              <th rowSpan={2} className="th-sticky left">รหัสพนักงาน</th>
              <th rowSpan={2} className="th-sticky">ช่วงขาดการติดต่อ</th>
              <th rowSpan={2} className="th-sticky right">งานทั้งหมด</th>

              <th colSpan={2} className="th-sticky group done">ทำได้</th>
              <th colSpan={2} className="th-sticky group self">มาเอง</th>
              <th colSpan={2} className="th-sticky group unpaid">ไม่จ่าย</th>

              <th rowSpan={2} className="th-sticky center">100%</th>
            </tr>
            <tr>
              <th className="th-sticky right done">จำนวน</th>
              <th className="th-sticky right done">%</th>

              <th className="th-sticky right self">จำนวน</th>
              <th className="th-sticky right self">%</th>

              <th className="th-sticky right unpaid">จำนวน</th>
              <th className="th-sticky right unpaid">%</th>
            </tr>
          </thead>

          <tbody>
            {rows.map((r) => (
              <tr
                key={r.key}
                className={[
                  r.isGrand ? 'row-grand' : '',
                  r.isSubtotal ? 'row-subtotal' : '',
                ].join(' ')}
              >
                <td className="cell left mono">{r.employeeId || '(ไม่ระบุ)'}</td>
                <td className="cell">{r.bucket}</td>
                <td className="cell right mono">{r.total.toLocaleString()}</td>

                <td className="cell right mono col-done">{r.done.toLocaleString()}</td>
                <td className="cell right mono col-done">{(r.pctDone * 100).toFixed(2)}%</td>

                <td className="cell right mono col-self">{r.selfPaid.toLocaleString()}</td>
                <td className="cell right mono col-self">{(r.pctSelf * 100).toFixed(2)}%</td>

                <td className="cell right mono col-unpaid">{r.unpaid.toLocaleString()}</td>
                <td className="cell right mono col-unpaid">{(r.pctUnpaid * 100).toFixed(2)}%</td>

                <td className="cell center mono">100.00%</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── Main Page ─────────────────────────────────────────────────────────────────

export default function Home() {
  const [file1Buffer, setFile1Buffer] = useState<ArrayBuffer | null>(null);
  const [file2Buffer, setFile2Buffer] = useState<ArrayBuffer | null>(null);
  const [file1Name, setFile1Name] = useState('');
  const [file2Name, setFile2Name] = useState('');

  const [isProcessing, setIsProcessing] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [result, setResult] = useState<ProcessingResult | null>(null);

  const [sortKey, setSortKey] = useState<SortKey>('index');
  const [sortDir, setSortDir] = useState<SortDir>('asc');
  const [statusFilter, setStatusFilter] = useState<StatusFilter>('ALL');
  const [bucketFilter, setBucketFilter] = useState('ALL');
  const [empFilter, setEmpFilter] = useState('ALL');

  // ── Handlers ────────────────────────────────────────────────────────────────

  const handleFile = useCallback((
    e: React.ChangeEvent<HTMLInputElement>,
    setBuffer: (b: ArrayBuffer) => void,
    setName: (s: string) => void,
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const buf = evt.target?.result;
      if (buf instanceof ArrayBuffer) setBuffer(buf);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleProcess = useCallback(() => {
    if (!file1Buffer || !file2Buffer) return;
    setIsProcessing(true);
    setTimeout(() => {
      try {
        const res = processExcelData(file1Buffer, file2Buffer);
        setResult(res);
        setSortKey('index');
        setSortDir('asc');
        setStatusFilter('ALL');
        setBucketFilter('ALL');
        setEmpFilter('ALL');
      } catch (err) {
        console.error(err);
        alert('เกิดข้อผิดพลาดในการประมวลผลไฟล์ กรุณาตรวจสอบหัวคอลัมน์ว่าตรงกับที่กำหนด');
      } finally {
        setIsProcessing(false);
      }
    }, 60);
  }, [file1Buffer, file2Buffer]);

  const handleSort = useCallback((key: SortKey) => {
    if (sortKey === key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(key);
      setSortDir('asc');
    }
  }, [sortKey]);

  const handleDownload = useCallback(async () => {
    if (!result) return;
    setIsDownloading(true);
    try {
      await generateOneExcelFile(result);
    } catch (err) {
      console.error(err);
      alert('เกิดข้อผิดพลาดในการสร้างไฟล์ Excel');
    } finally {
      setIsDownloading(false);
    }
  }, [result]);

  // ── Derived data ────────────────────────────────────────────────────────────

  const employeeIds = useMemo(() => {
    if (!result) return [];
    return ['ALL', ...Array.from(new Set(result.records.map((r) => r.employeeId).filter(Boolean))).sort()];
  }, [result]);

  const filteredSorted = useMemo(() => {
    if (!result) return [];
    let rows = result.records;
    if (statusFilter !== 'ALL') rows = rows.filter((r) => r.status === statusFilter);
    if (bucketFilter !== 'ALL') rows = rows.filter((r) => r.overdueBucket === bucketFilter);
    if (empFilter !== 'ALL')    rows = rows.filter((r) => r.employeeId === empFilter);

    return [...rows].sort((a, b) => {
      const av = a[sortKey];
      const bv = b[sortKey];
      const cmp = typeof av === 'number' && typeof bv === 'number'
        ? av - bv
        : String(av ?? '').localeCompare(String(bv ?? ''), 'th');
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }, [result, statusFilter, bucketFilter, empFilter, sortKey, sortDir]);

  // ── Helpers ─────────────────────────────────────────────────────────────────

  const Th = ({ label, col, align = 'left' }: { label: string; col: SortKey; align?: 'left' | 'right' | 'center' }) => (
    <th className={`th th-${align}`} onClick={() => handleSort(col)}>
      <span className="th-inner">
        {label}
        <SortIcon active={sortKey === col} dir={sortDir} />
      </span>
    </th>
  );

  return (
    <div className="page">
      <header className="header">
        <div className="header-inner">
          <div className="header-logo">📊</div>
          <div>
            <h1 className="header-title">Collection Performance</h1>
          </div>
        </div>
      </header>

      <main className="main">
        <section className="section">
          <h2 className="section-title">📁 อัปโหลดไฟล์</h2>
          <div className="upload-grid">
            <UploadZone
              label="1 · Collection File (ผลการตาม)"
              fileName={file1Name}
              accentClass="accent-blue"
              onChange={(e) => handleFile(e, setFile1Buffer, setFile1Name)}
            />
            <UploadZone
              label="2 · Export File (Transaction)"
              fileName={file2Name}
              accentClass="accent-emerald"
              onChange={(e) => handleFile(e, setFile2Buffer, setFile2Name)}
            />
          </div>

          <div className="btn-row">
            <button
              className="btn-process"
              onClick={handleProcess}
              disabled={!file1Buffer || !file2Buffer || isProcessing}
            >
              {isProcessing ? '⏳ กำลังประมวลผล…' : '⚡ ประมวลผลไฟล์'}
            </button>
          </div>
        </section>

        {result && (
          <div className="results">
            <section className="section">
              <h2 className="section-title">📈 สรุปผลการติดตาม</h2>
              <div className="summary-grid">
                <SummaryCard value={result.summary.total}    label="รายชื่อทั้งหมด"  colorClass="card-neutral" />
                <SummaryCard value={result.summary.selfPaid} label="จ่ายเอง"          colorClass="card-blue"    />
                <SummaryCard value={result.summary.callPaid} label="โทรตามมาจ่าย"    colorClass="card-green"   />
                <SummaryCard value={result.summary.unpaid}   label="ไม่จ่าย"          colorClass="card-red"     />
              </div>
            </section>

            {/* ✅ Summary Report Table */}
            <section className="section">
              <div className="table-header">
                <h2 className="section-title">📊 รายงานแบบตาราง (เหมือนรูป)</h2>
              </div>
              <SummaryReportTable result={result} />
            </section>

            {/* Detail Table */}
            <section className="section">
              <div className="table-header">
                <h2 className="section-title">📋 รายละเอียดสัญญา</h2>
                <div className="table-actions">
                  <span className="record-count">
                    แสดง {filteredSorted.length.toLocaleString()} / {result.records.length.toLocaleString()} รายการ
                  </span>
                  <button
                    className="btn-download"
                    onClick={handleDownload}
                    disabled={isDownloading}
                  >
                    {isDownloading ? '⏳ กำลังสร้างไฟล์…' : '⬇ ดาวน์โหลด output.xlsx (2 Sheets)'}
                  </button>
                </div>
              </div>

              <div className="filters">
                <div className="filter-group">
                  <label className="filter-label">ผลการตาม</label>
                  <div className="status-tabs">
                    {(
                      [
                        { key: 'ALL',       label: 'ทั้งหมด'       },
                        { key: 'SELF_PAID', label: 'จ่ายเอง'       },
                        { key: 'CALL_PAID', label: 'โทรตามมาจ่าย' },
                        { key: 'UNPAID',    label: 'ไม่จ่าย'        },
                      ] as { key: StatusFilter; label: string }[]
                    ).map(({ key, label }) => (
                      <button
                        key={key}
                        className={`status-tab ${statusFilter === key ? 'tab-active' : ''} ${
                          key === 'SELF_PAID' ? 'tab-blue'    :
                          key === 'CALL_PAID' ? 'tab-green'   :
                          key === 'UNPAID'    ? 'tab-red'     : 'tab-neutral'
                        }`}
                        onClick={() => setStatusFilter(key)}
                      >
                        {label}
                        {key !== 'ALL' && result && (
                          <span className="tab-count">
                            {key === 'SELF_PAID' ? result.summary.selfPaid :
                             key === 'CALL_PAID' ? result.summary.callPaid :
                             result.summary.unpaid}
                          </span>
                        )}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="filter-group">
                  <label className="filter-label">ช่วงขาดการติดต่อ</label>
                  <select
                    className="filter-select"
                    value={bucketFilter}
                    onChange={(e) => setBucketFilter(e.target.value)}
                  >
                    <option value="ALL">ทุกช่วง</option>
                    {OVERDUE_BUCKETS.map((b) => (
                      <option key={b} value={b}>{b}</option>
                    ))}
                  </select>
                </div>

                <div className="filter-group">
                  <label className="filter-label">รหัสพนักงาน</label>
                  <select
                    className="filter-select"
                    value={empFilter}
                    onChange={(e) => setEmpFilter(e.target.value)}
                  >
                    {employeeIds.map((id) => (
                      <option key={id} value={id}>{id === 'ALL' ? 'ทุกพนักงาน' : id}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="table-wrapper">
                <table className="data-table">
                  <thead>
                    <tr>
                      <Th label="ลำดับ"               col="index"          align="center" />
                      <Th label="รหัสพนักงาน"         col="employeeId"                    />
                      <Th label="สาขา"                col="branch"                        />
                      <Th label="เลขที่สัญญา"         col="contractNo"                    />
                      <Th label="ชื่อ-นามสกุล"        col="fullName"                      />
                      <Th label="วันขาดการติดต่อ"     col="overdueDays"    align="center" />
                      <Th label="ช่วงขาดการติดต่อ"    col="overdueBucket"                 />
                      <Th label="ยอดชำระรวม"         col="totalPaidAmount" align="right" />
                      <Th label="ผลการตาม"            col="status"         align="center" />
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSorted.length === 0 ? (
                      <tr>
                        <td colSpan={9} className="empty-row">ไม่พบข้อมูลที่ตรงกับเงื่อนไข</td>
                      </tr>
                    ) : (
                      filteredSorted.map((row, idx) => (
                        <tr key={idx} className={`data-row row-${row.status.toLowerCase()}`}>
                          <td className="td td-center">{row.index}</td>
                          <td className="td td-mono">{row.employeeId || '(ไม่ระบุ)'}</td>
                          <td className="td">{row.branch}</td>
                          <td className="td td-mono td-bold">{row.contractNo}</td>
                          <td className="td">{row.fullName}</td>
                          <td className="td td-center">{row.overdueDays}</td>
                          <td className="td">
                            <span className="bucket-pill">{row.overdueBucket}</span>
                          </td>
                          <td className="td td-right td-mono">
                            {row.totalPaidAmount > 0
                              ? `฿${row.totalPaidAmount.toLocaleString()}`
                              : <span className="td-dash">—</span>}
                          </td>
                          <td className="td td-center">
                            <StatusBadge status={row.status} />
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </div>
        )}
      </main>



      {/* ✅ CSS ทั้งหมดในไฟล์เดียว */}
      <style jsx global>{`
        :root{
          --bg: #070b14;
          --panel: rgba(255,255,255,0.04);
          --panel2: rgba(255,255,255,0.06);
          --stroke: rgba(255,255,255,0.10);
          --text: rgba(255,255,255,0.92);
          --muted: rgba(255,255,255,0.55);
          --muted2: rgba(255,255,255,0.35);
          --blue: #3b82f6;
          --green: #10b981;
          --red: #ef4444;
          --purple: #7c3aed;
          --doneCol: rgba(255, 209, 102, 0.16);
          --selfCol: rgba(16, 185, 129, 0.14);
          --unpaidCol: rgba(148, 163, 184, 0.18);
        }

        html, body { height: 100%; }
        body{
          margin:0;
          background: radial-gradient(1200px 600px at 10% 10%, rgba(124,58,237,0.10), transparent 60%),
                      radial-gradient(1200px 600px at 90% 0%, rgba(59,130,246,0.10), transparent 55%),
                      var(--bg);
          color: var(--text);
          font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, "Noto Sans Thai", "Helvetica Neue", Arial, "Apple Color Emoji", "Segoe UI Emoji";
        }

        .page{ min-height:100%; }
        .header{
          padding: 28px 20px 12px;
          border-bottom: 1px solid rgba(255,255,255,0.08);
          backdrop-filter: blur(10px);
          position: sticky; top: 0; z-index: 20;
          background: rgba(7,11,20,0.55);
        }
        .header-inner{ display:flex; gap:14px; align-items:center; max-width: 1200px; margin: 0 auto; }
        .header-logo{
          width:44px; height:44px; display:grid; place-items:center;
          border-radius: 14px;
          background: rgba(255,255,255,0.06);
          border: 1px solid rgba(255,255,255,0.12);
          box-shadow: 0 10px 30px rgba(0,0,0,0.35);
          font-size: 20px;
        }
        .header-title{ margin:0; font-size: 20px; letter-spacing: 0.2px; }
        .header-sub{ margin: 2px 0 0; color: var(--muted); font-size: 13px; }

        .main{ max-width: 1200px; margin: 0 auto; padding: 18px 20px 42px; }
        .section{
          background: var(--panel);
          border: 1px solid rgba(255,255,255,0.10);
          border-radius: 20px;
          padding: 18px;
          margin-bottom: 16px;
          box-shadow: 0 18px 60px rgba(0,0,0,0.35);
        }
        .section-title{ margin: 0 0 12px; font-size: 16px; }
        .upload-grid{ display:grid; grid-template-columns: 1fr 1fr; gap: 14px; }
        @media (max-width: 880px){ .upload-grid{ grid-template-columns: 1fr; } }

        .upload-card{
          border-radius: 18px;
          padding: 14px;
          background: rgba(255,255,255,0.03);
          border: 1px dashed rgba(255,255,255,0.16);
        }
        .accent-blue{ box-shadow: inset 0 0 0 1px rgba(59,130,246,0.25); }
        .accent-emerald{ box-shadow: inset 0 0 0 1px rgba(16,185,129,0.22); }
        .upload-label{ margin: 0 0 10px; color: var(--muted); font-weight: 600; font-size: 13px; }
        .upload-zone{
          display:block;
          cursor:pointer;
          border-radius: 16px;
          padding: 18px;
          background: rgba(255,255,255,0.03);
          border: 1px solid rgba(255,255,255,0.10);
        }
        .upload-idle{ display:flex; flex-direction:column; gap: 6px; align-items:center; color: var(--muted); }
        .upload-icon{ font-size: 22px; }
        .upload-hint{ font-size: 13px; }
        .upload-ext{ font-size: 12px; color: var(--muted2); }
        .upload-done{ display:flex; gap: 10px; align-items:center; color: rgba(255,255,255,0.9); }
        .upload-check{
          width: 22px; height: 22px; border-radius: 999px;
          background: rgba(16,185,129,0.18);
          border: 1px solid rgba(16,185,129,0.35);
          display:grid; place-items:center;
          color: #34d399;
          font-weight: 800;
        }
        .upload-filename{ font-size: 13px; word-break: break-all; }

        .btn-row{ display:flex; justify-content:center; margin-top: 14px; }
        .btn-process{
          border:0;
          border-radius: 999px;
          padding: 12px 18px;
          font-weight: 700;
          color: white;
          cursor:pointer;
          background: linear-gradient(135deg, rgba(124,58,237,1), rgba(59,130,246,1));
          box-shadow: 0 10px 30px rgba(124,58,237,0.25);
        }
        .btn-process:disabled{ opacity: 0.5; cursor:not-allowed; }

        .summary-grid{ display:grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }
        @media (max-width: 960px){ .summary-grid{ grid-template-columns: 1fr 1fr; } }
        @media (max-width: 520px){ .summary-grid{ grid-template-columns: 1fr; } }

        .summary-card{
          border-radius: 18px;
          padding: 16px;
          background: rgba(255,255,255,0.03);
          border: 1px solid rgba(255,255,255,0.10);
          position: relative;
          overflow:hidden;
        }
        .summary-card::before{
          content:'';
          position:absolute; left:0; top:0; bottom:0; width:4px;
          opacity: 0.9;
          background: rgba(255,255,255,0.15);
        }
        .card-blue::before{ background: rgba(59,130,246,1); }
        .card-green::before{ background: rgba(16,185,129,1); }
        .card-red::before{ background: rgba(239,68,68,1); }
        .card-neutral::before{ background: rgba(148,163,184,1); }

        .summary-value{ margin:0; font-size: 46px; font-weight: 800; letter-spacing: 0.5px; }
        .summary-label{ margin: 4px 0 0; color: var(--muted); font-weight: 600; }

        .table-header{ display:flex; align-items:center; justify-content:space-between; gap: 12px; margin-bottom: 12px; }
        .table-actions{ display:flex; align-items:center; gap: 10px; }
        .record-count{ color: var(--muted); font-size: 12px; }
        .btn-download{
          border:0;
          border-radius: 12px;
          padding: 10px 14px;
          cursor:pointer;
          font-weight: 700;
          color: white;
          background: rgba(16,185,129,0.22);
          border: 1px solid rgba(16,185,129,0.45);
        }
        .btn-download:disabled{ opacity: 0.55; cursor:not-allowed; }

        /* ✅ Summary Report table styles */
        .report-wrap{
          border-radius: 18px;
          background: rgba(255,255,255,0.02);
          border: 1px solid rgba(255,255,255,0.10);
          overflow:hidden;
        }
        .report-title{
          padding: 14px 14px 10px;
          border-bottom: 1px solid rgba(255,255,255,0.08);
        }
        .report-title-main{ font-weight: 800; font-size: 16px; }
        .report-title-sub{ margin-top: 4px; color: var(--muted); font-size: 12px; }

        .report-table-scroll{
          overflow:auto;
          max-height: 420px; /* เลื่อนลงได้ ถ้าข้อมูลเยอะ */
        }
        .report-table{
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          min-width: 980px; /* ให้เป็น table จริง + scroll แนวนอน */
          table-layout: fixed;
        }
        .report-table thead th{
          position: sticky;
          top: 0;
          z-index: 2;
          background: rgba(10,14,24,0.95);
          border-bottom: 1px solid rgba(255,255,255,0.10);
          font-size: 12px;
          color: rgba(255,255,255,0.85);
          padding: 10px 10px;
          text-align: center;
          white-space: nowrap;
        }
        .report-table thead .group{
          font-weight: 800;
          letter-spacing: 0.2px;
        }
        .report-table thead .done{ background: rgba(255, 209, 102, 0.10); }
        .report-table thead .self{ background: rgba(16,185,129,0.10); }
        .report-table thead .unpaid{ background: rgba(148,163,184,0.12); }

        .report-table tbody td{
          padding: 10px 10px;
          border-bottom: 1px solid rgba(255,255,255,0.06);
          font-size: 14px;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        .report-table tbody tr:hover td{
          background: rgba(255,255,255,0.03);
        }

        .cell.left{ text-align:left; }
        .cell.right{ text-align:right; }
        .cell.center{ text-align:center; }
        .mono{ font-variant-numeric: tabular-nums; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace; }

        .col-done{ background: var(--doneCol); }
        .col-self{ background: var(--selfCol); }
        .col-unpaid{ background: var(--unpaidCol); }

        .row-subtotal td{
          background: rgba(146, 208, 80, 0.10);
          font-weight: 800;
        }
        .row-grand td{
          background: rgba(255,255,255,0.06);
          font-weight: 900;
        }

        /* Detail table styles (your existing look) */
        .filters{ display:flex; flex-wrap:wrap; gap: 12px; margin: 10px 0 12px; }
        .filter-group{ display:flex; flex-direction:column; gap: 6px; }
        .filter-label{ color: var(--muted); font-size: 12px; font-weight: 700; }
        .filter-select{
          background: rgba(255,255,255,0.03);
          border: 1px solid rgba(255,255,255,0.10);
          color: var(--text);
          border-radius: 12px;
          padding: 10px 12px;
          min-width: 220px;
          outline: none;
        }
        .status-tabs{ display:flex; gap: 10px; flex-wrap:wrap; }
        .status-tab{
          cursor:pointer;
          padding: 10px 12px;
          border-radius: 999px;
          border: 1px solid rgba(255,255,255,0.12);
          background: rgba(255,255,255,0.02);
          color: rgba(255,255,255,0.85);
          display:flex; align-items:center; gap: 8px;
          font-weight: 800;
        }
        .tab-count{
          padding: 3px 8px;
          border-radius: 999px;
          background: rgba(255,255,255,0.06);
          border: 1px solid rgba(255,255,255,0.10);
          font-size: 12px;
          color: rgba(255,255,255,0.8);
        }
        .tab-active{ box-shadow: inset 0 0 0 1px rgba(255,255,255,0.22); }
        .tab-blue{ border-color: rgba(59,130,246,0.35); }
        .tab-green{ border-color: rgba(16,185,129,0.35); }
        .tab-red{ border-color: rgba(239,68,68,0.35); }

        .table-wrapper{ overflow:auto; border-radius: 16px; border: 1px solid rgba(255,255,255,0.10); }
        .data-table{
          width:100%;
          border-collapse: collapse;
          min-width: 980px;
        }
        .data-table thead th{
          position: sticky;
          top: 0;
          z-index: 1;
          background: rgba(10,14,24,0.95);
          border-bottom: 1px solid rgba(255,255,255,0.10);
          color: rgba(255,255,255,0.72);
          font-size: 12px;
          padding: 12px 10px;
          text-align:left;
          cursor:pointer;
          white-space: nowrap;
        }
        .th-inner{ display:flex; align-items:center; gap: 8px; }
        .sort-icon{ font-size: 12px; opacity: 0.65; }
        .sort-icon.active{ opacity: 1; }
        .sort-icon.inactive{ opacity: 0.5; }

        .td{
          padding: 12px 10px;
          border-bottom: 1px solid rgba(255,255,255,0.06);
          font-size: 13px;
          white-space: nowrap;
        }
        .td-center{ text-align:center; }
        .td-right{ text-align:right; }
        .td-mono{ font-variant-numeric: tabular-nums; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace; }
        .td-bold{ font-weight: 900; }
        .td-dash{ color: var(--muted2); }

        .bucket-pill{
          padding: 6px 10px;
          border-radius: 999px;
          background: rgba(124,58,237,0.14);
          border: 1px solid rgba(124,58,237,0.25);
          color: rgba(196,181,253,0.95);
          font-weight: 800;
          font-size: 12px;
        }

        .badge{
          display:inline-flex; align-items:center; gap: 8px;
          padding: 8px 10px;
          border-radius: 999px;
          font-weight: 900;
          font-size: 12px;
          border: 1px solid rgba(255,255,255,0.14);
        }
        .dot{ width: 8px; height: 8px; border-radius: 999px; }
        .badge-blue{ background: rgba(59,130,246,0.12); color: rgba(147,197,253,0.95); border-color: rgba(59,130,246,0.28); }
        .dot-blue{ background: rgba(59,130,246,1); }

        .badge-green{ background: rgba(16,185,129,0.12); color: rgba(110,231,183,0.95); border-color: rgba(16,185,129,0.28); }
        .dot-green{ background: rgba(16,185,129,1); }

        .badge-red{ background: rgba(239,68,68,0.12); color: rgba(252,165,165,0.95); border-color: rgba(239,68,68,0.28); }
        .dot-red{ background: rgba(239,68,68,1); }

        .empty-row{ padding: 20px; text-align:center; color: var(--muted); }

        .footer{
          padding: 20px;
          text-align:center;
          color: var(--muted2);
          font-size: 12px;
        }
      `}</style>
    </div>
  );
}