"use client";

import { Fragment, useEffect, useMemo, useState } from "react";
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
type WorkOrderStatus = "Planned" | "Released" | "In Process" | "Closed" | "Unknown";
type WorkOrderStatusRow = {
  workOrder: string;
  status: WorkOrderStatus;
  item: string;
  itemKey: string;
  description: string;
  qtyNeeded: number;
  prodStartDate: string;
};
type WorkOrderMaterialRow = {
  workOrder: string;
  item: string;
  itemKey: string;
  description: string;
  qtyNeeded: number;
  qtyOnHand: number;
  qtyIssued: number;
  units: string;
};
type WorkOrderPartLine = WorkOrderMaterialRow & {
  status: WorkOrderStatus;
  assemblyItem: string;
  assemblyDescription: string;
};
type WorkOrderPartContext = {
  openDemand: number;
  issuedQty: number;
  inProcessCount: number;
  plannedReleasedCount: number;
  closedCount: number;
  activeWoCount: number;
  lines: WorkOrderPartLine[];
};
type SortDir = "asc" | "desc";
type SummarySortKey = "date" | "unitPn" | "unitDescription" | "pwbPn" | "requiredQty" | "unitCount" | "openDemand" | "issuedQty";
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

function normWo(v: string): string {
  return cellText(v).trim().toUpperCase();
}

function itemKey(v: string): string {
  return normPart(v).split(":")[0].trim();
}

function parseWoStatus(v: string): WorkOrderStatus {
  const text = cellText(v).toLowerCase();
  if (text === "planned") return "Planned";
  if (text === "released") return "Released";
  if (text === "in process" || text === "in-process" || text === "inprocess") return "In Process";
  if (text === "closed") return "Closed";
  return "Unknown";
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

    const rawHeaders = rows[headerIdx] || [];
    const headers = rawHeaders.map((c) => normalizeHeader(cellText(c)));
    const productIdx = findHeaderIndex(headers, [/^category$/, /^product$/, /^program$/]);
    const itemIdx = findHeaderIndex(headers, [/^item$/, /^item_number$/, /^part_number$/]);
    const descIdx = findHeaderIndex(headers, [/^description$/, /^desc$/]);
    const avgUpIdx = findHeaderIndex(headers, [/^unit_price$/, /^avg_up$/, /^avg$/, /^price$/]);
    const onHandIdx = findHeaderIndex(headers, [/^on_hand$/, /^onhand$/]);

    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const product = productIdx >= 0 ? cellText(row[productIdx]) : "";
      const unitPn = normPart(cellText(row[itemIdx >= 0 ? itemIdx : 1]));
      if (!unitPn) continue;
      const description = descIdx >= 0 ? cellText(row[descIdx]) : "";
      const avgUp = avgUpIdx >= 0 ? cellText(row[avgUpIdx]) : "";
      const onHand = onHandIdx >= 0 ? cellText(row[onHandIdx]) : "";
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

function parseWorkOrderStatusWorkbook(wb: XLSX.WorkBook): WorkOrderStatusRow[] {
  const out: WorkOrderStatusRow[] = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = worksheetRows(ws).filter((r) => r.some((c) => cellText(c)));
    if (rows.length < 2) continue;

    let headerIdx = -1;
    let woIdx = -1;
    let statusIdx = -1;
    let itemIdx = -1;
    let descIdx = -1;
    let qtyIdx = -1;
    let startIdx = -1;
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const headers = rows[r].map((c) => normalizeHeader(cellText(c)));
      const w = findHeaderIndex(headers, [/^work_order$/, /^wo$/]);
      const s = findHeaderIndex(headers, [/^status$/]);
      const i = findHeaderIndex(headers, [/^item$/, /^assembly$/, /^part_number$/]);
      if (w >= 0 && s >= 0 && i >= 0) {
        headerIdx = r;
        woIdx = w;
        statusIdx = s;
        itemIdx = i;
        descIdx = findHeaderIndex(headers, [/^description$/, /^desc$/]);
        qtyIdx = findHeaderIndex(headers, [/^qty_needed$/, /^quantity_needed$/, /^qty$/]);
        startIdx = findHeaderIndex(headers, [/^prod_start_date$/, /^production_start_date$/, /^start_date$/]);
        break;
      }
    }
    if (headerIdx < 0) continue;

    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      const workOrder = normWo(cellText(row[woIdx]));
      const item = normPart(cellText(row[itemIdx]));
      if (!workOrder || !item) continue;
      out.push({
        workOrder,
        status: parseWoStatus(cellText(row[statusIdx])),
        item,
        itemKey: itemKey(item),
        description: descIdx >= 0 ? cellText(row[descIdx]) : "",
        qtyNeeded: qtyIdx >= 0 ? cellNumber(row[qtyIdx]) : 0,
        prodStartDate: startIdx >= 0 ? cellText(row[startIdx]) : "",
      });
    }
  }

  const dedup = new Map<string, WorkOrderStatusRow>();
  for (const row of out) dedup.set(row.workOrder, row);
  return Array.from(dedup.values()).sort((a, b) => a.workOrder.localeCompare(b.workOrder));
}

function parseWorkOrderMaterialWorkbook(wb: XLSX.WorkBook): WorkOrderMaterialRow[] {
  const out: WorkOrderMaterialRow[] = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = worksheetRows(ws).filter((r) => r.some((c) => cellText(c)));
    if (rows.length < 2) continue;

    let headerIdx = -1;
    let woIdx = -1;
    let itemIdx = -1;
    let descIdx = -1;
    let qtyNeededIdx = -1;
    let qtyOnHandIdx = -1;
    let qtyIssuedIdx = -1;
    let unitsIdx = -1;
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const headers = rows[r].map((c) => normalizeHeader(cellText(c)));
      const w = findHeaderIndex(headers, [/^work_order$/, /^wo$/]);
      const i = findHeaderIndex(headers, [/^item$/, /^part_number$/, /^component$/]);
      const needed = findHeaderIndex(headers, [/^qty_needed$/, /^quantity_needed$/]);
      const issued = findHeaderIndex(headers, [/^qty_issued$/, /^quantity_issued$/]);
      if (w >= 0 && i >= 0 && needed >= 0 && issued >= 0) {
        headerIdx = r;
        woIdx = w;
        itemIdx = i;
        qtyNeededIdx = needed;
        qtyIssuedIdx = issued;
        descIdx = findHeaderIndex(headers, [/^description$/, /^desc$/]);
        qtyOnHandIdx = findHeaderIndex(headers, [/^qty_on_hand$/, /^on_hand$/, /^qty_available$/]);
        unitsIdx = findHeaderIndex(headers, [/^units$/, /^unit$/]);
        break;
      }
    }
    if (headerIdx < 0) continue;

    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      const workOrder = normWo(cellText(row[woIdx]));
      const item = normPart(cellText(row[itemIdx]));
      if (!workOrder || !item) continue;
      out.push({
        workOrder,
        item,
        itemKey: itemKey(item),
        description: descIdx >= 0 ? cellText(row[descIdx]) : "",
        qtyNeeded: cellNumber(row[qtyNeededIdx]),
        qtyOnHand: qtyOnHandIdx >= 0 ? cellNumber(row[qtyOnHandIdx]) : 0,
        qtyIssued: cellNumber(row[qtyIssuedIdx]),
        units: unitsIdx >= 0 ? cellText(row[unitsIdx]) : "",
      });
    }
  }

  return out.sort((a, b) => a.workOrder.localeCompare(b.workOrder) || a.item.localeCompare(b.item));
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
  const [workOrderStatusFile, setWorkOrderStatusFile] = useState("");
  const [workOrderMaterialFile, setWorkOrderMaterialFile] = useState("");
  const [forecastRows, setForecastRows] = useState<ForecastDemand[]>([]);
  const [mappingRows, setMappingRows] = useState<UnitPwbMapRow[]>([]);
  const [workOrderStatusRows, setWorkOrderStatusRows] = useState<WorkOrderStatusRow[]>([]);
  const [workOrderMaterialRows, setWorkOrderMaterialRows] = useState<WorkOrderMaterialRow[]>([]);
  const [boms, setBoms] = useState<BomItem[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [filterToKnownPwbs, setFilterToKnownPwbs] = useState(true);
  const [expandedPwb, setExpandedPwb] = useState<string | null>(null);
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

  async function loadWorkOrderStatus(file: File) {
    setError(null);
    setWorkOrderStatusFile(file.name);
    try {
      const wb = await readWorkbook(file);
      const rows = parseWorkOrderStatusWorkbook(wb);
      setWorkOrderStatusRows(rows);
      if (rows.length === 0) setError("No WO status rows found. Expected Work Order, Status, Item, and Qty Needed columns.");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to parse WO status report");
    }
  }

  async function loadWorkOrderMaterial(file: File) {
    setError(null);
    setWorkOrderMaterialFile(file.name);
    try {
      const wb = await readWorkbook(file);
      const rows = parseWorkOrderMaterialWorkbook(wb);
      setWorkOrderMaterialRows(rows);
      if (rows.length === 0) setError("No WO material-issued rows found. Expected Work Order, Item, Qty Needed, and Qty Issued columns.");
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Failed to parse WO material-issued report");
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

  const workOrderStatusByWo = useMemo(() => {
    const map = new Map<string, WorkOrderStatusRow>();
    for (const row of workOrderStatusRows) map.set(row.workOrder, row);
    return map;
  }, [workOrderStatusRows]);

  const workOrderContextByPart = useMemo(() => {
    const map = new Map<string, WorkOrderPartContext>();
    const getContext = (key: string) => {
      let context = map.get(key);
      if (!context) {
        context = { openDemand: 0, issuedQty: 0, inProcessCount: 0, plannedReleasedCount: 0, closedCount: 0, activeWoCount: 0, lines: [] };
        map.set(key, context);
      }
      return context;
    };
    const seenWoByPart = new Map<string, Set<string>>();

    for (const row of workOrderMaterialRows) {
      if (!row.itemKey || row.workOrder.startsWith("R")) continue;
      const active = workOrderStatusByWo.get(row.workOrder);
      const status = active?.status || "Closed";
      const context = getContext(row.itemKey);
      context.lines.push({
        ...row,
        status,
        assemblyItem: active?.item || "",
        assemblyDescription: active?.description || "",
      });

      const seenKey = row.itemKey;
      if (!seenWoByPart.has(seenKey)) seenWoByPart.set(seenKey, new Set());
      const seen = seenWoByPart.get(seenKey)!;
      const firstWoPartLine = !seen.has(row.workOrder);
      seen.add(row.workOrder);

      if (status === "Planned" || status === "Released") {
        context.openDemand += row.qtyNeeded;
        if (firstWoPartLine) {
          context.activeWoCount += 1;
          context.plannedReleasedCount += 1;
        }
      } else if (status === "In Process") {
        context.issuedQty += row.qtyIssued;
        if (firstWoPartLine) {
          context.activeWoCount += 1;
          context.inProcessCount += 1;
        }
      } else if (status === "Closed") {
        if (firstWoPartLine) context.closedCount += 1;
      }
    }

    for (const context of map.values()) {
      context.lines.sort((a, b) => {
        const statusOrder = { "In Process": 0, Released: 1, Planned: 2, Unknown: 3, Closed: 4 } as Record<WorkOrderStatus, number>;
        return statusOrder[a.status] - statusOrder[b.status] || a.workOrder.localeCompare(b.workOrder) || b.qtyIssued - a.qtyIssued || b.qtyNeeded - a.qtyNeeded;
      });
    }
    return map;
  }, [workOrderMaterialRows, workOrderStatusByWo]);

  const workOrderStats = useMemo(() => {
    const activeNonRework = workOrderStatusRows.filter((r) => !r.workOrder.startsWith("R"));
    const materialWos = new Set(workOrderMaterialRows.map((r) => r.workOrder).filter((wo) => !wo.startsWith("R")));
    let inferredClosed = 0;
    for (const wo of materialWos) if (!workOrderStatusByWo.has(wo)) inferredClosed += 1;
    return {
      active: activeNonRework.length,
      material: materialWos.size,
      inferredClosed,
      inProcess: activeNonRework.filter((r) => r.status === "In Process").length,
      plannedReleased: activeNonRework.filter((r) => r.status === "Planned" || r.status === "Released").length,
    };
  }, [workOrderMaterialRows, workOrderStatusByWo, workOrderStatusRows]);

  const summary = useMemo(() => {
    const agg = new Map<string, { date: string; pwbPn: string; requiredQty: number; unitCount: number; unitPns: Set<string>; unitDescriptions: Set<string> }>();
    for (const r of pwbDemand) {
      const key = `${r.date}||${r.pwbPn}`;
      const ex = agg.get(key);
      if (ex) {
        ex.requiredQty += r.requiredQty;
        ex.unitCount += 1;
        ex.unitPns.add(r.unitPn);
        if (r.unitDescription) ex.unitDescriptions.add(r.unitDescription);
      } else {
        agg.set(key, {
          date: r.date,
          pwbPn: r.pwbPn,
          requiredQty: r.requiredQty,
          unitCount: 1,
          unitPns: new Set([r.unitPn]),
          unitDescriptions: new Set(r.unitDescription ? [r.unitDescription] : []),
        });
      }
    }
    const rows = Array.from(agg.values()).map((r) => {
      const woContext = workOrderContextByPart.get(itemKey(r.pwbPn)) || { openDemand: 0, issuedQty: 0, inProcessCount: 0, plannedReleasedCount: 0, closedCount: 0, activeWoCount: 0, lines: [] };
      return {
        date: r.date,
        pwbPn: r.pwbPn,
        requiredQty: r.requiredQty,
        unitCount: r.unitCount,
        unitPn: Array.from(r.unitPns).sort().join("; "),
        unitDescription: Array.from(r.unitDescriptions).sort().join("; "),
        woContext,
      };
    });
    rows.sort((a, b) => {
      let cmp = 0;
      switch (summarySort.key) {
        case "date": cmp = a.date.localeCompare(b.date); break;
        case "unitPn": cmp = a.unitPn.localeCompare(b.unitPn); break;
        case "unitDescription": cmp = a.unitDescription.localeCompare(b.unitDescription); break;
        case "pwbPn": cmp = a.pwbPn.localeCompare(b.pwbPn); break;
        case "requiredQty": cmp = a.requiredQty - b.requiredQty; break;
        case "unitCount": cmp = a.unitCount - b.unitCount; break;
        case "openDemand": cmp = a.woContext.openDemand - b.woContext.openDemand; break;
        case "issuedQty": cmp = a.woContext.issuedQty - b.woContext.issuedQty; break;
      }
      if (cmp === 0) cmp = a.date.localeCompare(b.date) || a.pwbPn.localeCompare(b.pwbPn);
      return summarySort.dir === "asc" ? cmp : -cmp;
    });
    return rows;
  }, [pwbDemand, summarySort, workOrderContextByPart]);

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

  function woCoverageLabel(context: WorkOrderPartContext) {
    const labels: string[] = [];
    if (context.inProcessCount) labels.push(`${context.inProcessCount} In Process`);
    if (context.plannedReleasedCount) labels.push(`${context.plannedReleasedCount} Planned/Released`);
    if (context.closedCount) labels.push(`${context.closedCount} Closed`);
    return labels.length > 0 ? labels.join(", ") : "No WO match";
  }

  function exportDetail() {
    downloadCsv("iop-pwb-demand-detail.csv", [
      ["date", "product", "unit_pn", "unit_description", "unit_qty", "pwb_pn", "qty_per_unit", "required_qty", "wo_open_demand", "wo_issued_qty", "wo_coverage"],
      ...pwbDemand.map((r) => {
        const context = workOrderContextByPart.get(itemKey(r.pwbPn)) || { openDemand: 0, issuedQty: 0, inProcessCount: 0, plannedReleasedCount: 0, closedCount: 0, activeWoCount: 0, lines: [] };
        return [r.date, r.product, r.unitPn, r.unitDescription, r.unitQty, r.pwbPn, r.qtyPerUnit, r.requiredQty, context.openDemand, context.issuedQty, woCoverageLabel(context)];
      }),
    ]);
  }

  function exportSummary() {
    downloadCsv("iop-pwb-demand-summary.csv", [
      ["date", "unit_level_part_number", "unit_level_part_number_description", "pwb_pn", "required_qty", "source_unit_lines", "wo_open_demand", "wo_issued_qty", "wo_coverage"],
      ...summary.map((r) => [r.date, r.unitPn, r.unitDescription, r.pwbPn, r.requiredQty, r.unitCount, r.woContext.openDemand, r.woContext.issuedQty, woCoverageLabel(r.woContext)]),
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

      <section className="border rounded p-4 space-y-3">
        <div>
          <h2 className="font-semibold">3) WO context reports</h2>
          <p className="text-xs text-gray-500">Upload WO status plus WO Material Issued. Planned/Released counts as open demand; In Process counts as already issued; missing from active status but present in material history is inferred Closed. R* rework WOs are ignored here.</p>
        </div>
        <div className="grid md:grid-cols-2 gap-4">
          <div className="space-y-2">
            <div className="text-xs font-medium">WO Status</div>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => { const f = e.target.files?.[0]; if (f) void loadWorkOrderStatus(f); }} />
            {workOrderStatusFile && <div className="text-xs text-gray-600">Loaded: {workOrderStatusFile}</div>}
            <div className="text-xs">Active non-rework WOs: <b>{workOrderStats.active}</b></div>
          </div>
          <div className="space-y-2">
            <div className="text-xs font-medium">WO Material Issued</div>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => { const f = e.target.files?.[0]; if (f) void loadWorkOrderMaterial(f); }} />
            {workOrderMaterialFile && <div className="text-xs text-gray-600">Loaded: {workOrderMaterialFile}</div>}
            <div className="text-xs">Non-rework WOs in material history: <b>{workOrderStats.material}</b></div>
          </div>
        </div>
      </section>

      {error && <div className="border border-red-300 bg-red-50 text-red-700 rounded p-3">{error}</div>}

      <section className="grid md:grid-cols-4 gap-3">
        <div className="border rounded p-3">Forecast rows: <b>{forecastRows.length}</b></div>
        <div className="border rounded p-3">Unit/PWB mappings: <b>{mappingRows.length}</b></div>
        <div className="border rounded p-3">PWB demand lines: <b>{pwbDemand.length}</b></div>
        <div className="border rounded p-3">Unmapped Unit PNs: <b>{missingUnits.length}</b></div>
        <div className="border rounded p-3">WO Planned/Released: <b>{workOrderStats.plannedReleased}</b></div>
        <div className="border rounded p-3">WO In Process: <b>{workOrderStats.inProcess}</b></div>
        <div className="border rounded p-3">Inferred Closed WOs: <b>{workOrderStats.inferredClosed}</b></div>
        <div className="border rounded p-3">PWB WO matches: <b>{summary.filter((r) => r.woContext.lines.length > 0).length}</b></div>
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
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("unitPn")}>Unit Level Part Number{summarySortLabel("unitPn")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("unitDescription")}>Unit Level Part Number Description{summarySortLabel("unitDescription")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("pwbPn")}>PWB{summarySortLabel("pwbPn")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("requiredQty")}>Required Qty{summarySortLabel("requiredQty")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("openDemand")}>WO Open Demand{summarySortLabel("openDemand")}</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("issuedQty")}>Issued to WO{summarySortLabel("issuedQty")}</th>
                  <th className="p-2">WO Coverage</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleSummarySort("unitCount")}>Source Lines{summarySortLabel("unitCount")}</th>
                </tr>
              </thead>
              <tbody>
                {summary.map((r) => {
                  const rowKey = `${r.date}-${r.pwbPn}`;
                  const expanded = expandedPwb === rowKey;
                  return (
                    <Fragment key={rowKey}>
                      <tr className="border-b">
                        <td className="p-2 whitespace-nowrap">{r.date}</td>
                        <td className="p-2 font-mono">{r.unitPn}</td>
                        <td className="p-2 min-w-72">{r.unitDescription || "—"}</td>
                        <td className="p-2 font-mono">{r.pwbPn}</td>
                        <td className="p-2 font-semibold">{r.requiredQty}</td>
                        <td className="p-2 font-semibold text-amber-700">{r.woContext.openDemand || "—"}</td>
                        <td className="p-2 font-semibold text-blue-700">{r.woContext.issuedQty || "—"}</td>
                        <td className="p-2 min-w-48">
                          <button
                            className={`text-left ${r.woContext.lines.length > 0 ? "underline decoration-dotted" : "text-gray-500"}`}
                            disabled={r.woContext.lines.length === 0}
                            onClick={() => setExpandedPwb((prev) => prev === rowKey ? null : rowKey)}
                          >
                            {woCoverageLabel(r.woContext)}
                          </button>
                        </td>
                        <td className="p-2">{r.unitCount}</td>
                      </tr>
                      {expanded && (
                        <tr className="border-b bg-gray-50">
                          <td className="p-3" colSpan={9}>
                            <div className="mb-2 text-xs font-semibold">WO detail for {r.pwbPn}</div>
                            <div className="overflow-auto max-h-80 rounded border bg-white">
                              <table className="w-full text-xs">
                                <thead className="sticky top-0 bg-white">
                                  <tr className="text-left border-b">
                                    <th className="p-2">WO</th>
                                    <th className="p-2">Status</th>
                                    <th className="p-2">Assembly / Sub-assembly</th>
                                    <th className="p-2">Material</th>
                                    <th className="p-2">Qty Needed</th>
                                    <th className="p-2">Qty Issued</th>
                                    <th className="p-2">On Hand</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {r.woContext.lines.map((line, idx) => (
                                    <tr key={`${line.workOrder}-${line.item}-${idx}`} className="border-b">
                                      <td className="p-2 font-mono">{line.workOrder}</td>
                                      <td className="p-2">{line.status}</td>
                                      <td className="p-2 font-mono">{line.assemblyItem || "—"}{line.assemblyDescription ? <div className="font-sans text-gray-500">{line.assemblyDescription}</div> : null}</td>
                                      <td className="p-2 font-mono">{line.item}<div className="font-sans text-gray-500">{line.description || "—"}</div></td>
                                      <td className="p-2 font-semibold">{line.qtyNeeded || "—"}</td>
                                      <td className="p-2 font-semibold text-blue-700">{line.qtyIssued || "—"}</td>
                                      <td className="p-2">{line.qtyOnHand || "—"}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  );
                })}
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
