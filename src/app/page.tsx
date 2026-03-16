"use client";

import { useCallback, useMemo, useState } from "react";
import {
  processExcelData,
  generateOneExcelFile,
  ProcessingResult,
  CollectionRecord,
  OVERDUE_BUCKETS,
} from "../utils/excelProcessor";

// ─── Types ─────────────────────────────────────────────────────────────────────

type SortKey = keyof CollectionRecord;
type SortDir = "asc" | "desc";
type StatusFilter = "ALL" | "SELF_PAID" | "CALL_PAID" | "UNPAID";

// ─── Status config ──────────────────────────────────────────────────────────────

const STATUS_CONFIG = {
  SELF_PAID: { label: "จ่ายเอง", badge: "badge-blue", dot: "dot-blue" },
  CALL_PAID: { label: "โทรตามมาจ่าย", badge: "badge-green", dot: "dot-green" },
  UNPAID: { label: "ไม่จ่าย", badge: "badge-red", dot: "dot-red" },
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
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={onChange}
          className="sr-only"
        />
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

function SummaryCard({
  value,
  label,
  colorClass,
}: {
  value: number;
  label: string;
  colorClass: string;
}) {
  return (
    <div className={`summary-card ${colorClass}`}>
      <p className="summary-value">{value.toLocaleString()}</p>
      <p className="summary-label">{label}</p>
    </div>
  );
}

function StatusBadge({ status }: { status: CollectionRecord["status"] }) {
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
  return <span className="sort-icon active">{dir === "asc" ? "↑" : "↓"}</span>;
}

type SummaryRow = {
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
};

type PerformanceRow = {
  key: string;
  employeeId: string;
  bucket: string;
  baseTotal: number;
  done: number;
  pctDone: number;
  unpaid: number;
  pctUnpaid: number;
  isSubtotal: boolean;
  isGrand: boolean;
};

function buildSummaryRows(result: ProcessingResult): SummaryRow[] {
  const out: SummaryRow[] = [];

  for (const emp of result.employeeSummary) {
    for (const b of emp.buckets) {
      if (b.total === 0) continue;

      out.push({
        key: `${emp.employeeId}-${b.bucket}`,
        employeeId: emp.employeeId,
        bucket: b.bucket,
        total: b.total,
        done: b.callPaid,
        pctDone: b.total ? b.callPaid / b.total : 0,
        selfPaid: b.selfPaid,
        pctSelf: b.total ? b.selfPaid / b.total : 0,
        unpaid: b.unpaid,
        pctUnpaid: b.total ? b.unpaid / b.total : 0,
        isSubtotal: false,
        isGrand: false,
      });
    }

    out.push({
      key: `${emp.employeeId}-subtotal`,
      employeeId: emp.employeeId,
      bucket: "รวม",
      total: emp.total,
      done: emp.callPaid,
      pctDone: emp.total ? emp.callPaid / emp.total : 0,
      selfPaid: emp.selfPaid,
      pctSelf: emp.total ? emp.selfPaid / emp.total : 0,
      unpaid: emp.unpaid,
      pctUnpaid: emp.total ? emp.unpaid / emp.total : 0,
      isSubtotal: true,
      isGrand: false,
    });
  }

  return out;
}

function buildPerformanceRows(result: ProcessingResult): PerformanceRow[] {
  const out: PerformanceRow[] = [];

  for (const emp of result.employeeSummary) {
    for (const b of emp.buckets) {
      if (b.total === 0) continue;
      const baseTotal = b.total - b.selfPaid;

      out.push({
        key: `${emp.employeeId}-${b.bucket}-performance`,
        employeeId: emp.employeeId,
        bucket: b.bucket,
        baseTotal,
        done: b.callPaid,
        pctDone: baseTotal ? b.callPaid / baseTotal : 0,
        unpaid: b.unpaid,
        pctUnpaid: baseTotal ? b.unpaid / baseTotal : 0,
        isSubtotal: false,
        isGrand: false,
      });
    }

    const baseTotal = emp.total - emp.selfPaid;
    out.push({
      key: `${emp.employeeId}-subtotal-performance`,
      employeeId: emp.employeeId,
      bucket: "รวม",
      baseTotal,
      done: emp.callPaid,
      pctDone: baseTotal ? emp.callPaid / baseTotal : 0,
      unpaid: emp.unpaid,
      pctUnpaid: baseTotal ? emp.unpaid / baseTotal : 0,
      isSubtotal: true,
      isGrand: false,
    });
  }

  return out;
}

// ─── Summary Table (HTML table จริง) ───────────────────────────────────────────

function SummaryReportTable({ result }: { result: ProcessingResult }) {
  const rows = useMemo(() => buildSummaryRows(result), [result]);
  const performanceRows = useMemo(() => buildPerformanceRows(result), [result]);

  return (
    <div className="report-wrap">
      <div className="report-title">
        <div className="report-title-main">{result.reportTitle}</div>
      </div>
      <div className="report-table-scroll">
        <table className="report-table">
          <thead>
            <tr>
              <th rowSpan={2} className="th-sticky left">
                รหัสพนักงาน
              </th>
              <th rowSpan={2} className="th-sticky">
                ช่วงขาดการติดต่อ
              </th>
              <th rowSpan={2} className="th-sticky right">
                งานทั้งหมด
              </th>

              <th colSpan={2} className="th-sticky group done">
                ทำได้
              </th>
              <th colSpan={2} className="th-sticky group self">
                ไม่จ่าย
              </th>
              <th colSpan={2} className="th-sticky group unpaid">
                ลูกค้ามาจ่ายเอง
              </th>

              <th rowSpan={2} className="th-sticky center">
                100%
              </th>
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
                  r.isGrand ? "row-grand" : "",
                  r.isSubtotal ? "row-subtotal" : "",
                ].join(" ")}
              >
                <td className="cell center mono">{r.employeeId || "-"}</td>
                <td className="cell center">{r.bucket}</td>
                <td className="cell center mono">{r.total.toLocaleString()}</td>

                <td className="cell center mono col-done">
                  {r.done.toLocaleString()}
                </td>
                <td className="cell center mono col-done">
                  {(r.pctDone * 100).toFixed(2)}%
                </td>

                <td className="cell center mono col-unpaid">
                  {r.unpaid.toLocaleString()}
                </td>
                <td className="cell center mono col-unpaid">
                  {(r.pctUnpaid * 100).toFixed(2)}%
                </td>

                <td className="cell center mono col-self">
                  {r.selfPaid.toLocaleString()}
                </td>
                <td className="cell center mono col-self">
                  {(r.pctSelf * 100).toFixed(2)}%
                </td>

                <td className="cell center mono">100.00%</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="report-title report-title-secondary">
        <div className="report-title-main">ผลการติดตามความสำเสร็จ</div>
      </div>
      <div className="report-table-scroll report-table-scroll-secondary">
        <table className="report-table report-table-secondary">
          <thead>
            <tr>
              <th rowSpan={2} className="th-sticky left">
                รหัสพนักงาน
              </th>
              <th rowSpan={2} className="th-sticky">
                ช่วงขาดการติดต่อ
              </th>
              <th rowSpan={2} className="th-sticky center">
                งานที่ตามเองทั้งหมด
              </th>
              <th colSpan={2} className="th-sticky group done">
                ทำได้
              </th>
              <th colSpan={2} className="th-sticky group self">
                ตามจ่ายไม่ได้
              </th>
            </tr>
            <tr>
              <th className="th-sticky center done">จำนวน</th>
              <th className="th-sticky center done">%</th>
              <th className="th-sticky center self">จำนวน</th>
              <th className="th-sticky center self">%</th>
            </tr>
          </thead>
          <tbody>
            {performanceRows.map((r) => (
              <tr
                key={r.key}
                className={[
                  r.isGrand ? "row-grand" : "",
                  r.isSubtotal ? "row-subtotal" : "",
                ].join(" ")}
              >
                <td className="cell center mono">{r.employeeId || "-"}</td>
                <td className="cell center">{r.bucket}</td>
                <td className="cell center mono">
                  {r.baseTotal.toLocaleString()}
                </td>
                <td className="cell center mono col-done">
                  {r.done.toLocaleString()}
                </td>
                <td className="cell center mono col-done">
                  {(r.pctDone * 100).toFixed(2)}%
                </td>
                <td className="cell center mono col-unpaid">
                  {r.unpaid.toLocaleString()}
                </td>
                <td className="cell center mono col-unpaid">
                  {(r.pctUnpaid * 100).toFixed(2)}%
                </td>
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
  const [file1Name, setFile1Name] = useState("");
  const [file2Name, setFile2Name] = useState("");

  const [isProcessing, setIsProcessing] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [result, setResult] = useState<ProcessingResult | null>(null);

  const [sortKey, setSortKey] = useState<SortKey>("index");
  const [sortDir, setSortDir] = useState<SortDir>("asc");
  const [statusFilter, setStatusFilter] = useState<StatusFilter>("ALL");
  const [bucketFilter, setBucketFilter] = useState("ALL");
  const [empFilter, setEmpFilter] = useState("ALL");

  // ── Handlers ────────────────────────────────────────────────────────────────

  const handleFile = useCallback(
    (
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
    },
    [],
  );

  const handleProcess = useCallback(() => {
    if (!file1Buffer || !file2Buffer) return;
    setIsProcessing(true);
    setTimeout(() => {
      try {
        const res = processExcelData(file1Buffer, file2Buffer);
        setResult(res);
        setSortKey("index");
        setSortDir("asc");
        setStatusFilter("ALL");
        setBucketFilter("ALL");
        setEmpFilter("ALL");
      } catch (err) {
        console.error(err);
        alert(
          "เกิดข้อผิดพลาดในการประมวลผลไฟล์ กรุณาตรวจสอบหัวคอลัมน์ว่าตรงกับที่กำหนด",
        );
      } finally {
        setIsProcessing(false);
      }
    }, 60);
  }, [file1Buffer, file2Buffer]);

  const handleSort = useCallback(
    (key: SortKey) => {
      if (sortKey === key) {
        setSortDir((d) => (d === "asc" ? "desc" : "asc"));
      } else {
        setSortKey(key);
        setSortDir("asc");
      }
    },
    [sortKey],
  );

  const handleDownload = useCallback(async () => {
    if (!result) return;
    setIsDownloading(true);
    try {
      await generateOneExcelFile(result);
    } catch (err) {
      console.error(err);
      alert("เกิดข้อผิดพลาดในการสร้างไฟล์ Excel");
    } finally {
      setIsDownloading(false);
    }
  }, [result]);

  // ── Derived data ────────────────────────────────────────────────────────────

  const filteredSorted = useMemo(() => {
    if (!result) return [];
    let rows = result.records;
    if (statusFilter !== "ALL")
      rows = rows.filter((r) => r.status === statusFilter);
    if (bucketFilter !== "ALL")
      rows = rows.filter((r) => r.overdueBucket === bucketFilter);
    if (empFilter !== "ALL")
      rows = rows.filter((r) => r.employeeId === empFilter);

    return [...rows].sort((a, b) => {
      const av = a[sortKey];
      const bv = b[sortKey];
      const cmp =
        typeof av === "number" && typeof bv === "number"
          ? av - bv
          : String(av ?? "").localeCompare(String(bv ?? ""), "th");
      return sortDir === "asc" ? cmp : -cmp;
    });
  }, [result, statusFilter, bucketFilter, empFilter, sortKey, sortDir]);

  // ── Helpers ─────────────────────────────────────────────────────────────────

  const Th = ({
    label,
    col,
    align = "left",
  }: {
    label: string;
    col: SortKey;
    align?: "left" | "right" | "center";
  }) => (
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
              {isProcessing ? "⏳ กำลังประมวลผล…" : "⚡ ประมวลผลไฟล์"}
            </button>
          </div>
        </section>

        {result && (
          <div className="results">
            <section className="section">
              <h2 className="section-title">📈 สรุปผลการติดตาม</h2>
              <div className="summary-grid">
                <SummaryCard
                  value={result.summary.total}
                  label="รายชื่อทั้งหมด"
                  colorClass="card-neutral"
                />
                <SummaryCard
                  value={result.summary.selfPaid}
                  label="จ่ายเอง"
                  colorClass="card-blue"
                />
                <SummaryCard
                  value={result.summary.callPaid}
                  label="โทรตามมาจ่าย"
                  colorClass="card-green"
                />
                <SummaryCard
                  value={result.summary.unpaid}
                  label="ไม่จ่าย"
                  colorClass="card-red"
                />
              </div>
            </section>

            {/* ✅ Summary Report Table */}
            <section className="section">
              <div className="table-header">
                <h2 className="section-title">📊 สรุปผลการตาม</h2>
              </div>
              <SummaryReportTable result={result} />
            </section>

            {/* Detail Table */}
            <section className="section">
              <div className="table-header">
                <h2 className="section-title">📋 รายละผลการตาม</h2>
                <div className="table-actions">
                  <span className="record-count">
                    แสดง {filteredSorted.length.toLocaleString()} /{" "}
                    {result.records.length.toLocaleString()} รายการ
                  </span>
                  <button
                    className="btn-download"
                    onClick={handleDownload}
                    disabled={isDownloading}
                  >
                    {isDownloading
                      ? "⏳ กำลังสร้างไฟล์…"
                      : `⬇ ดาวน์โหลด ${result.reportTitle}.xlsx (3 Sheets)`}
                  </button>
                </div>
              </div>

              <div className="filters">
                <div className="filter-group">
                  <label className="filter-label">ผลการตาม</label>
                  <div className="status-tabs">
                    {(
                      [
                        { key: "ALL", label: "ทั้งหมด" },
                        { key: "SELF_PAID", label: "จ่ายเอง" },
                        { key: "CALL_PAID", label: "โทรตามมาจ่าย" },
                        { key: "UNPAID", label: "ไม่จ่าย" },
                      ] as { key: StatusFilter; label: string }[]
                    ).map(({ key, label }) => (
                      <button
                        key={key}
                        className={`status-tab ${statusFilter === key ? "tab-active" : ""} ${
                          key === "SELF_PAID"
                            ? "tab-blue"
                            : key === "CALL_PAID"
                              ? "tab-green"
                              : key === "UNPAID"
                                ? "tab-red"
                                : "tab-neutral"
                        }`}
                        onClick={() => setStatusFilter(key)}
                      >
                        {label}
                        {key !== "ALL" && result && (
                          <span className="tab-count">
                            {key === "SELF_PAID"
                              ? result.summary.selfPaid
                              : key === "CALL_PAID"
                                ? result.summary.callPaid
                                : result.summary.unpaid}
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
                      <option key={b} value={b}>
                        {b}
                      </option>
                    ))}
                  </select>
                </div>

                {/* <div className="filter-group">
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
                </div> */}
              </div>

              <div className="table-wrapper">
                <table className="data-table">
                  <thead>
                    <tr>
                      <Th label="ลำดับ" col="index" align="center" />
                      <Th label="รหัสพนักงาน" col="employeeId" />
                      <Th label="สาขา" col="branch" />
                      <Th label="เลขที่สัญญา" col="contractNo" />
                      <Th label="ชื่อ-นามสกุล" col="fullName" />
                      <Th
                        label="วันขาดการติดต่อ"
                        col="overdueDays"
                        align="center"
                      />
                      <Th label="ช่วงขาดการติดต่อ" col="overdueBucket" />
                      <Th
                        label="ยอดชำระรวม"
                        col="totalPaidAmount"
                        align="right"
                      />
                      <Th label="ผลการตาม" col="status" align="center" />
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSorted.length === 0 ? (
                      <tr>
                        <td colSpan={9} className="empty-row">
                          ไม่พบข้อมูลที่ตรงกับเงื่อนไข
                        </td>
                      </tr>
                    ) : (
                      filteredSorted.map((row, idx) => (
                        <tr
                          key={idx}
                          className={`data-row row-${row.status.toLowerCase()}`}
                        >
                          <td className="td td-center">{row.index}</td>
                          <td className="td td-mono">
                            {row.employeeId || "-"}
                          </td>
                          <td className="td">{row.branch}</td>
                          <td className="td td-mono td-bold">
                            {row.contractNo}
                          </td>
                          <td className="td">{row.fullName}</td>
                          <td className="td td-center">{row.overdueDays}</td>
                          <td className="td">
                            <span className="bucket-pill">
                              {row.overdueBucket}
                            </span>
                          </td>
                          <td className="td td-right td-mono">
                            {row.totalPaidAmount > 0 ? (
                              `฿${row.totalPaidAmount.toLocaleString()}`
                            ) : (
                              <span className="td-dash">—</span>
                            )}
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
    </div>
  );
}
