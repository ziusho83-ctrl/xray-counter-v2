"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Cell = string | number | boolean | Date | null | undefined;
type BomItem = { assembly_pn: string; revision: string; line_count: number; bom_type?: string };
type ForecastDemand = {
  product: string;
  unitPn: string;
  description: string;
  avgUp: string;
  onHand: string;
  date: string;
  qty: number;
  sheet: string;
  sourceFields: Record<string, string>;
};
type UnitPwbMapRow = {
  unitPn: string;
  pwbPn: string;
  qtyPerUnit: number;
  sourceSheet: string;
};
type PwbDemand = {
  date: string;
  product: string;
  unitPn: string;
  unitDescription: string;
  unitQty: number;
  pwbPn: string;
  qtyPerUnit: number;
  requiredQty: number;
};
type SortDir = "asc" | "desc";
type SummarySortKey = "date" | "pwbPn" | "requiredQty" | "unitCount";
type MissingSortKey = "unitPn" | "description" | "product" | "qty";

const MONTHS: Record<string, number> = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
};

function cellText(v: Cell): string {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v).trim();
}

function cellNumber(v: Cell): number {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const text = cellText(v).replace(/[$,]/g, "").trim();
  if (!text || text === "-" || text === "–") return 0;
  const n = Number(text);
  return Number.isFinite(n) ? n : 0;
}

function normPart(v: string): string {
  return v.trim().toUpperCase();
}

function inferYear(fileName: string, sheetNames: string[]): number {
  const hay = `${fileName} ${sheetNames.join(" ")}`;
  const full = hay.match(/20\d{2}/);
  if (full) return Number(full[0]);
  const yy = hay.match(/(?:^|[^\d])(\d{2})(?:$|[^\d])/);
  if (yy) return 2000 + Number(yy[1]);
  return new Date().getFullYear();
}

function parseDateHeader(v: Cell, fallbackYear: number): string | null {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    if (v < 20000) return null;
    const parsed = XLSX.SSF.parse_date_code(v);
    if (parsed?.y && parsed?.m && parsed?.d) {
      return new Date(parsed.y, parsed.m - 1, parsed.d).toISOString().slice(0, 10);
    }
  }
  const text = cellText(v);
  if (!text) return null;

  const mdy = text.match(/^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$/);
  if (mdy) {
    const month = Number(mdy[1]) - 1;
    const day = Number(mdy[2]);
    const year = mdy[3] ? (Number(mdy[3]) < 100 ? 2000 + Number(mdy[3]) : Number(mdy[3])) : fallbackYear;
    return new Date(year, month, day).toISOString().slice(0, 10);
  }

  const dayMonth = text.match(/^(\d{1,2})[-\s]?([A-Za-z]{3,9})(?:[-\s]?(\d{2,4}))?$/);
  if (dayMonth) {
    const day = Number(dayMonth[1]);
    const month = MONTHS[dayMonth[2].toLowerCase()];
    if (month === undefined) return null;
    const year = dayMonth[3] ? (Number(dayMonth[3]) < 100 ? 2000 + Number(dayMonth[3]) : Number(dayMonth[3])) : fallbackYear;
    return new Date(year, month, day).toISOString().slice(0, 10);
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return null;
}

function worksheetRows(ws: XLSX.WorkSheet): Cell[][] {
  return XLSX.utils.sheet_to_json<Cell[]>(ws, { header: 1, defval: null, raw: true });
}

function parseForecastWorkbook(wb: XLSX.WorkBook, fileName: string): ForecastDemand[] {
  const fallbackYear = inferYear(fileName, wb.SheetNames);
  const demand: ForecastDemand[] = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = worksheetRows(ws);
    if (rows.length < 2) continue;

    const salesOrderDemand = parseSalesOrderDemandSheet(rows, sheetName, fallbackYear);
    if (salesOrderDemand) {
      demand.push(...salesOrderDemand);
      continue;
    }

    let headerIdx = -1;
    let dateCols: Array<{ col: number; date: string }> = [];
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r] || [];
      const first = cellText(row[0]).toLowerCase();
      const second = cellText(row[1]).toLowerCase();
      const cols: Array<{ col: number; date: string }> = [];
      for (let c = 2; c < row.length; c++) {
        const date = parseDateHeader(row[c], fallbackYear);
        if (date) cols.push({ col: c, date });
      }
      if ((first.includes("product") || second.includes("item")) && cols.length > 0) {
        headerIdx = r;
        dateCols = cols;
        break;
      }
      if (cols.length >= 3 && headerIdx < 0) {
        headerIdx = r;
        dateCols = cols;
      }
    }
    if (headerIdx < 0 || dateCols.length === 0) continue;

    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const product = cellText(row[0]);
      const unitPn = normPart(cellText(row[1]));
      if (!unitPn) continue;
      const description = cellText(row[2]);
      const avgUp = cellText(row[3]);
      const onHand = cellText(row[4]);
      const sourceFields = {
        PRODUCT: product,
        ITEM: unitPn,
        DESC: description,
        "AVG UP": avgUp,
        "ON HAND": onHand,
      };
      for (const dc of dateCols) {
        const qty = cellNumber(row[dc.col]);
        if (qty > 0) {
          demand.push({ product, unitPn, description, avgUp, onHand, date: dc.date, qty, sheet: sheetName, sourceFields });
        }
      }
    }
  }

  return demand.sort((a, b) => a.date.localeCompare(b.date) || a.unitPn.localeCompare(b.unitPn));
}

function normalizeHeader(v: string): string {
  return v.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function findHeaderIndex(headers: string[], patterns: RegExp[]): number {
  return headers.findIndex((h) => patterns.some((p) => p.test(h)));
}

function sourceFieldsFromRow(headers: string[], row: Cell[]): Record<string, string> {
  const out: Record<string, string> = {};
  headers.forEach((h, i) => {
    const header = h.trim();
    if (!header) return;
    out[header] = cellText(row[i]);
  });
  return out;
}

function parseSalesOrderDemandSheet(rows: Cell[][], sheetName: string, fallbackYear: number): ForecastDemand[] | null {
  for (let r = 0; r < Math.min(rows.length, 10); r++) {
    const rawHeaders = (rows[r] || []).map((c) => cellText(c));
    const headers = rawHeaders.map(normalizeHeader);
    const itemIdx = findHeaderIndex(headers, [/^item$/, /^item_number$/, /^part_number$/]);
    const qtyIdx = findHeaderIndex(headers, [/^quantity$/, /^qty$/, /^qty_required$/, /^q$/]);
    const dateIdx = findHeaderIndex(headers, [/^mps_date$/, /^mps$/, /^current_promise_date$/, /^customer_contract_date$/, /^demand_date$/]);
    if (itemIdx < 0 || qtyIdx < 0 || dateIdx < 0) continue;

    const productIdx = findHeaderIndex(headers, [/^sales_order_number$/, /^so$/, /^product$/, /^customer$/]);
    const descIdx = findHeaderIndex(headers, [/^memo$/, /^desc$/, /^description$/]);
    const out: ForecastDemand[] = [];
    for (let i = r + 1; i < rows.length; i++) {
      const row = rows[i] || [];
      const unitPn = normPart(cellText(row[itemIdx]));
      const qty = cellNumber(row[qtyIdx]);
      const date = parseDateHeader(row[dateIdx], fallbackYear);
      if (!unitPn || qty <= 0 || !date) continue;
      const description = descIdx >= 0 ? cellText(row[descIdx]) : "";
      out.push({
        product: productIdx >= 0 ? cellText(row[productIdx]) : "",
        unitPn,
        description,
        avgUp: "",
        onHand: "",
        date,
        qty,
        sheet: sheetName,
        sourceFields: sourceFieldsFromRow(rawHeaders, row),
      });
    }
    return out;
  }
  return null;
}

function parseUnitPwbWorkbook(wb: XLSX.WorkBook, pwbSet: Set<string>, filterToKnownPwbs: boolean): UnitPwbMapRow[] {
  const out: UnitPwbMapRow[] = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = worksheetRows(ws).filter((r) => r.some((c) => cellText(c)));
    if (rows.length < 2) continue;

    let headerIdx = 0;
    let unitIdx = -1;
    let pwbIdx = -1;
    let qtyIdx = -1;
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const headers = rows[r].map((c) => normalizeHeader(cellText(c)));
      const u = findHeaderIndex(headers, [/unit.*(pn|part|item)/, /parent.*(pn|part|item)/, /top.*(pn|part|item)/, /assembly.*(pn|part|item)/, /^unit$/]);
      const p = findHeaderIndex(headers, [/pwb/, /pwa/, /board/, /component.*(pn|part|item)/, /^component$/, /^item$/]);
      const q = findHeaderIndex(headers, [/qty.*per/, /quantity/, /^qty$/]);
      if (u >= 0 && p >= 0 && u !== p) {
        headerIdx = r;
        unitIdx = u;
        pwbIdx = p;
        qtyIdx = q;
        break;
      }
    }
    if (unitIdx < 0 || pwbIdx < 0) {
      headerIdx = 0;
      unitIdx = 0;
      pwbIdx = 1;
      qtyIdx = 2;
    }

    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      const unitPn = normPart(cellText(row[unitIdx]));
      const pwbPn = normPart(cellText(row[pwbIdx]));
      if (!unitPn || !pwbPn || unitPn === pwbPn) continue;
      const knownPwb = pwbSet.has(pwbPn);
      const looksPwb = /(^6F-|PWB|PWA)/i.test(pwbPn);
      if (filterToKnownPwbs && pwbSet.size > 0 && !knownPwb) continue;
      if (!filterToKnownPwbs && !knownPwb && !looksPwb) continue;
      const rawQty = qtyIdx >= 0 ? cellNumber(row[qtyIdx]) : 0;
      out.push({ unitPn, pwbPn, qtyPerUnit: rawQty > 0 ? rawQty : 1, sourceSheet: sheetName });
    }
  }

  const dedup = new Map<string, UnitPwbMapRow>();
  for (const r of out) {
    const key = `${r.unitPn}||${r.pwbPn}`;
    const ex = dedup.get(key);
    if (ex) ex.qtyPerUnit += r.qtyPerUnit;
    else dedup.set(key, { ...r });
  }
  return Array.from(dedup.values()).sort((a, b) => a.unitPn.localeCompare(b.unitPn) || a.pwbPn.localeCompare(b.pwbPn));
}

function downloadCsv(filename: string, rows: Array<Array<string | number>>) {
  const cell = (v: string | number) => '"' + String(v ?? "").replace(/"/g, '""') + '"';
  const csv = rows.map((r) => r.map(cell).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export default function IopPage() {
  const [forecastFile, setForecastFile] = useState("");
  const [mappingFile, setMappingFile] = useState("");
  const [forecastRows, setForecastRows] = useState<ForecastDemand[]>([]);
  const [mappingRows, setMappingRows] = useState<UnitPwbMapRow[]>([]);
  const [boms, setBoms] = useState<BomItem[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [filterToKnownPwbs, setFilterToKnownPwbs] = useState(true);
  const [summarySort, setSummarySort] = useState<{ key: SummarySortKey; dir: SortDir }>({ key: "date", dir: "asc" });
  const [missingSort, setMissingSort] = useState<{ key: MissingSortKey; dir: SortDir }>({ key: "qty", dir: "desc" });

  useEffect(() => {
    void (async () => {
      try {
        const res = await fetch("/api/boms");
        const data = await res.json();
        setBoms(data.items || []);
      } catch { /* silent */ }
    })();
  }, []);

  const pwbSet = useMemo(() => new Set(boms.filter((b) => (b.bom_type || "PWB") === "PWB").map((b) => normPart(b.assembly_pn))), [boms]);

  async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
    const buf = await file.arrayBuffer();
    return XLSX.read(buf, { cellDates: true });
  }

  async function loadForecast(file: File) {
    setError(null);
    setForecastFile(file.name);
    try {
      const wb = await readWorkbook(file);
      const rows = parseForecastWorkbook(wb, file.name);
      setForecastRows(rows);
      if (rows.length === 0) setError("No dated demand quantities found. Expected Product in A, Unit PN in B, dates across row 1, quantities below those date columns.");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to parse forecast workbook");
    }
  }

  async function loadMapping(file: File) {
    setError(null);
    setMappingFile(file.name);
    try {
      const wb = await readWorkbook(file);
      const rows = parseUnitPwbWorkbook(wb, pwbSet, filterToKnownPwbs);
      setMappingRows(rows);
      if (rows.length === 0) setError("No Unit → PWB rows found. Try turning off 'Known PWB only' if this mapping file uses new PWB numbers not yet in Data Manager.");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to parse Unit/PWB breakdown");
    }
  }

  const pwbDemand = useMemo<PwbDemand[]>(() => {
    const map = new Map<string, UnitPwbMapRow[]>();
    for (const m of mappingRows) {
      if (!map.has(m.unitPn)) map.set(m.unitPn, []);
      map.get(m.unitPn)!.push(m);
    }
    const rows: PwbDemand[] = [];
    for (const f of forecastRows) {
      const pwbs = map.get(f.unitPn) || [];
      for (const p of pwbs) {
        rows.push({
          date: f.date,
          product: f.product,
          unitPn: f.unitPn,
          unitDescription: f.description,
          unitQty: f.qty,
          pwbPn: p.pwbPn,
          qtyPerUnit: p.qtyPerUnit,
          requiredQty: f.qty * p.qtyPerUnit,
        });
      }
    }
    return rows.sort((a, b) => a.date.localeCompare(b.date) || a.pwbPn.localeCompare(b.pwbPn) || a.unitPn.localeCompare(b.unitPn));
  }, [forecastRows, mappingRows]);

  const summary = useMemo(() => {
    const agg = new Map<string, { date: string; pwbPn: string; requiredQty: number; unitCount: number }>();
    for (const r of pwbDemand) {
      const key = `${r.date}||${r.pwbPn}`;
      const ex = agg.get(key);
      if (ex) {
        ex.requiredQty += r.requiredQty;
        ex.unitCount += 1;
      } else {
        agg.set(key, { date: r.date, pwbPn: r.pwbPn, requiredQty: r.requiredQty, unitCount: 1 });
      }
    }
    const rows = Array.from(agg.values());
    rows.sort((a, b) => {
      let cmp = 0;
      switch (summarySort.key) {
        case "date": cmp = a.date.localeCompare(b.date); break;
        case "pwbPn": cmp = a.pwbPn.localeCompare(b.pwbPn); break;
        case "requiredQty": cmp = a.requiredQty - b.requiredQty; break;
        case "unitCount": cmp = a.unitCount - b.unitCount; break;
      }
      if (cmp === 0) cmp = a.date.localeCompare(b.date) || a.pwbPn.localeCompare(b.pwbPn);
      return summarySort.dir === "asc" ? cmp : -cmp;
    });
    return rows;
  }, [pwbDemand, summarySort]);

  const missingUnits = useMemo(() => {
    const mapped = new Set(mappingRows.map((m) => m.unitPn));
    const miss = new Map<string, { unitPn: string; description: string; product: string; qty: number }>();
    for (const f of forecastRows) {
      if (mapped.has(f.unitPn)) continue;
      const ex = miss.get(f.unitPn);
      if (ex) {
        ex.qty += f.qty;
        if (!ex.description && f.description) ex.description = f.description;
        if (!ex.product && f.product) ex.product = f.product;
      } else {
        miss.set(f.unitPn, { unitPn: f.unitPn, description: f.description, product: f.product, qty: f.qty });
      }
    }
    const rows = Array.from(miss.values());
    rows.sort((a, b) => {
      let cmp = 0;
      switch (missingSort.key) {
        case "unitPn": cmp = a.unitPn.localeCompare(b.unitPn); break;
        case "description": cmp = a.description.localeCompare(b.description); break;
        case "product": cmp = a.product.localeCompare(b.product); break;
        case "qty": cmp = a.qty - b.qty; break;
      }
      if (cmp === 0) cmp = a.unitPn.localeCompare(b.unitPn);
      return missingSort.dir === "asc" ? cmp : -cmp;
    });
    return rows;
  }, [forecastRows, mappingRows, missingSort]);

  const forecastPreviewColumns = useMemo(() => {
    const cols: string[] = [];
    const seen = new Set<string>();
    for (const r of forecastRows.slice(0, 200)) {
      for (const k of Object.keys(r.sourceFields)) {
        if (!seen.has(k)) {
          seen.add(k);
          cols.push(k);
        }
      }
    }
    return cols;
  }, [forecastRows]);

  function toggleSummarySort(key: SummarySortKey) {
    setSummarySort((prev) => prev.key === key ? { key, dir: prev.dir === "asc" ? "desc" : "asc" } : { key, dir: key === "requiredQty" || key === "unitCount" ? "desc" : "asc" });
  }

  function toggleMissingSort(key: MissingSortKey) {
    setMissingSort((prev) => prev.key === key ? { key, dir: prev.dir === "asc" ? "desc" : "asc" } : { key, dir: key === "qty" ? "desc" : "asc" });
  }

  function summarySortLabel(key: SummarySortKey) {
    return summarySort.key === key ? (summarySort.dir === "asc" ? " ↑" : " ↓") : " ↕";
  }

  function missingSortLabel(key: MissingSortKey) {
    return missingSort.key === key ? (missingSort.dir === "asc" ? " ↑" : " ↓") : " ↕";
  }

  function exportDetail() {
    downloadCsv("iop-pwb-demand-detail.csv", [
      ["date", "product", "unit_pn", "unit_description", "unit_qty", "pwb_pn", "qty_per_unit", "required_qty"],
      ...pwbDemand.map((r) => [r.date, r.product, r.unitPn, r.unitDescription, r.unitQty, r.pwbPn, r.qtyPerUnit, r.requiredQty]),
    ]);
  }

  function exportSummary() {
    downloadCsv("iop-pwb-demand-summary.csv", [
      ["date", "pwb_pn", "required_qty", "source_unit_lines"],
      ...summary.map((r) => [r.date, r.pwbPn, r.requiredQty, r.unitCount]),
    ]);
  }

  return (
    <main className="max-w-6xl mx-auto p-6 space-y-6 text-sm">
      <div>
        <h1 className="text-2xl font-bold">IOP</h1>
        <p className="text-gray-600 mt-1">Forecast Unit-level demand, map it to PWB assemblies, then list which PWBs are needed by date.</p>
      </div>

      <section className="grid md:grid-cols-2 gap-4">
        <div className="border rounded p-4 space-y-3">
          <h2 className="font-semibold">1) Demand forecast</h2>
          <p className="text-xs text-gray-500">Expected: Product in column A, Unit-level PN in column B, dates across row 1, quantities under date columns.</p>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => { const f = e.target.files?.[0]; if (f) void loadForecast(f); }} />
          {forecastFile && <div className="text-xs text-gray-600">Loaded: {forecastFile}</div>}
          <div className="text-xs">Parsed demand rows: <b>{forecastRows.length}</b></div>
        </div>

        <div className="border rounded p-4 space-y-3">
          <h2 className="font-semibold">2) Unit → PWB breakdown</h2>
          <p className="text-xs text-gray-500">Upload the BOM breakdown when ready. I’ll detect Unit/Parent PN, PWB/Board/Component PN, and optional qty-per-unit columns.</p>
          <label className="flex items-center gap-2 text-xs">
            <input
              type="checkbox"
              checked={filterToKnownPwbs}
              onChange={(e) => {
                setFilterToKnownPwbs(e.target.checked);
                setMappingRows([]);
                setMappingFile("");
              }}
            />
            Known PWB only ({pwbSet.size} PWB assemblies from Data Manager)
          </label>
          <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => { const f = e.target.files?.[0]; if (f) void loadMapping(f); }} />
          {mappingFile && <div className="text-xs text-gray-600">Loaded: {mappingFile}</div>}
          <div className="text-xs">Mapped Unit/PWB rows: <b>{mappingRows.length}</b></div>
        </div>
      </section>

      {error && <div className="border border-red-300 bg-red-50 text-red-700 rounded p-3">{error}</div>}

      <section className="grid md:grid-cols-4 gap-3">
        <div className="border rounded p-3">Forecast rows: <b>{forecastRows.length}</b></div>
        <div className="border rounded p-3">Unit/PWB mappings: <b>{mappingRows.length}</b></div>
        <div className="border rounded p-3">PWB demand lines: <b>{pwbDemand.length}</b></div>
        <div className="border rounded p-3">Unmapped Unit PNs: <b>{missingUnits.length}</b></div>
      </section>

      {summary.length > 0 && (
        <section className="border rounded p-4 space-y-3">
          <div className="flex items-center justify-between gap-3 flex-wrap">
            <h2 className="font-semibold">PWB Demand by Need Date</h2>
            <div className="flex gap-2">
              <button className="rounded border px-3 py-1" onClick={exportSummary}>Export Summary CSV</button>
              <button className="rounded border px-3 py-1" onClick={exportDetail}>Export Detail CSV</button>
            </div>
          </div>
          <div className="overflow-auto max-h-[520px]">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-white">
                <tr className="text-left border-b">
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("date")}>Need Date{summarySortLabel("date")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("pwbPn")}>PWB{summarySortLabel("pwbPn")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("requiredQty")}>Required Qty{summarySortLabel("requiredQty")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("unitCount")}>Source Lines{summarySortLabel("unitCount")}</th>
                </tr>
              </thead>
              <tbody>
                {summary.map((r) => (
                  <tr key={`${r.date}-${r.pwbPn}`} className="border-b">
                    <td className="p-2 whitespace-nowrap">{r.date}</td>
                    <td className="p-2 font-mono">{r.pwbPn}</td>
                    <td className="p-2 font-semibold">{r.requiredQty}</td>
                    <td className="p-2">{r.unitCount}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {forecastRows.length > 0 && missingUnits.length > 0 && (
        <section className="border border-amber-300 bg-amber-50 rounded p-4 space-y-2">
          <h2 className="font-semibold text-amber-900">Unit PNs waiting for PWB breakdown</h2>
          <p className="text-xs text-amber-800">These forecast items do not have Unit → PWB mapping yet. Click a column header to sort.</p>
          <div className="overflow-auto max-h-80 bg-white/60 rounded">
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-amber-50"><tr className="text-left border-b">
                <th className="p-2 cursor-pointer select-none hover:bg-amber-100" onClick={() => toggleMissingSort("unitPn")}>Unit PN{missingSortLabel("unitPn")}</th>
                <th className="p-2 cursor-pointer select-none hover:bg-amber-100" onClick={() => toggleMissingSort("description")}>Description{missingSortLabel("description")}</th>
                <th className="p-2 cursor-pointer select-none hover:bg-amber-100" onClick={() => toggleMissingSort("product")}>Product{missingSortLabel("product")}</th>
                <th className="p-2 cursor-pointer select-none hover:bg-amber-100" onClick={() => toggleMissingSort("qty")}>Total Forecast Qty{missingSortLabel("qty")}</th>
              </tr></thead>
              <tbody>{missingUnits.map((m) => <tr key={m.unitPn} className="border-b"><td className="p-2 font-mono">{m.unitPn}</td><td className="p-2">{m.description || "—"}</td><td className="p-2">{m.product || "—"}</td><td className="p-2 font-semibold">{m.qty}</td></tr>)}</tbody>
            </table>
          </div>
        </section>
      )}

      {forecastRows.length > 0 && (
        <section className="border rounded p-4 space-y-2">
          <h2 className="font-semibold">Forecast Preview</h2>
          <div className="overflow-auto max-h-72">
            <table className="w-full text-xs">
              <thead><tr className="text-left border-b">
                {forecastPreviewColumns.map((col) => <th key={col} className="p-2 whitespace-nowrap">{col}</th>)}
                <th className="p-2 whitespace-nowrap">DEMAND DATE</th>
                <th className="p-2 whitespace-nowrap">QTY REQUIREMENT</th>
                <th className="p-2 whitespace-nowrap">SOURCE SHEET</th>
              </tr></thead>
              <tbody>{forecastRows.slice(0, 200).map((r, i) => <tr key={`${r.sheet}-${r.unitPn}-${r.date}-${i}`} className="border-b">
                {forecastPreviewColumns.map((col) => <td key={col} className="p-2 whitespace-nowrap">{r.sourceFields[col] || ""}</td>)}
                <td className="p-2 whitespace-nowrap">{r.date}</td>
                <td className="p-2 font-semibold">{r.qty}</td>
                <td className="p-2">{r.sheet}</td>
              </tr>)}</tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  );
}
