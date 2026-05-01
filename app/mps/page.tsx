"use client";

import { Fragment, useCallback, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type BomItem = { assembly_pn: string; revision: string; line_count: number; bom_type?: string };

type MpsRow = {
  level: number;
  part: string;
  description: string;
  qty_req: number;
  avail: number;
  short: number;
  order: string;
  ord_qty: number;
  due_date: string;
  remarks: string;
  top_pn: string;
};

type MatchedBoard = {
  part: string;
  description: string;
  qty: number;
  short: number;
  order: string;
  revision: string;
  revisions: string[];
  priority: number;
  mps_level: number;
  top_pn: string;
  iop_next_demand: string;
};

type WoPhaseBadge = {
  label: string;
  className: string;
};

type StockroomResult = {
  part: string;
  shortage_qty: number;
  stockroom_available: number;
  confirmed_qty: number;
  inspection_qty: number;
  stockroom_locations: Array<{ bin_location: string; lot_number: string; qty: number }>;
  bom_run: string;
  mount_type?: "SMT" | "TH";
};

type ReadinessStatus = "clear-xray" | "clear-with-stockroom" | "not-clear";

type ReadinessRow = {
  run_id?: string;
  run_label?: string;
  assembly_pn: string;
  revision: string;
  description: string;
  qty: number;
  status: ReadinessStatus;
  max_buildable_xray: number;
  max_buildable_combined: number;
  shortage_parts_xray: number;
  shortage_parts_after_stockroom: number;
  details: Array<{
    part: string;
    required: number;
    xray_available: number;
    stockroom_available: number;
    combined_available: number;
    shortage: number;
    is_floor_stock: boolean;
    mount_type: "SMT" | "TH";
    is_shared_conflict: boolean;
    is_shared: boolean;
  }>;
};

type AnalysisResult = {
  source: string;
  summary: {
    requested_runs: number;
    requested_boards: number;
    individually_completable_runs: number;
    recommended_complete_runs: number;
    recommended_boards: number;
  };
  plans: Array<{
    run_id?: string;
    run_label?: string;
    assembly_pn: string;
    revision: string;
    qty: number;
    priority: number;
    can_complete: boolean;
    max_buildable: number;
    in_recommended: boolean;
    shortages: Array<{ part: string; required: number; available: number; shortage: number }>;
  }>;
  recommended: Array<{ run_id?: string; run_label?: string; assembly_pn: string; revision: string; qty: number; priority: number }>;
  top_shortages: Array<{ part: string; shortage: number; required: number; available: number; bom_run: string; priority: number }>;
  errors?: string[];
};

function csvCell(value: string | number | null | undefined): string {
  const text = String(value ?? "");
  return '"' + text.replace(/"/g, '""') + '"';
}

export default function MpsPage() {
  const [boms, setBoms] = useState<BomItem[]>([]);
  const [mpsRows, setMpsRows] = useState<MpsRow[]>([]);
  const [matched, setMatched] = useState<MatchedBoard[]>([]);
  const [woSmtStatusMap, setWoSmtStatusMap] = useState<Map<string, string>>(new Map());
  const [woThStatusMap, setWoThStatusMap] = useState<Map<string, string>>(new Map());
  const [woSmtStatusFileName, setWoSmtStatusFileName] = useState<string>("");
  const [woThStatusFileName, setWoThStatusFileName] = useState<string>("");
  const [openPOs, setOpenPOs] = useState<Array<{ vendor: string; po: string; line: string; item: string; description: string; qty_ordered: number; qty_received: number; qty_due: number; promise_date: string; original_promise: string }>>([]);
  const [fileNames, setFileNames] = useState<string[]>([]);
  const [parseError, setParseError] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const [workbooksRef, setWorkbooksRef] = useState<Array<{ name: string; wb: XLSX.WorkBook }>>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [shortageView, setShortageView] = useState<"by-bom" | "cross-bom">("by-bom");
  const [collapsedBoms, setCollapsedBoms] = useState<Set<string>>(new Set());
  const [stockroomResults, setStockroomResults] = useState<StockroomResult[] | null>(null);
  const [stockroomLoading, setStockroomLoading] = useState(false);
  const [stockroomError, setStockroomError] = useState<string | null>(null);
  const [stockroomSort, setStockroomSort] = useState<{ col: string; asc: boolean }>({ col: "shortage_qty", asc: false });
  const [floorStockSet, setFloorStockSet] = useState<Set<string>>(new Set());
  const [smtPartsSet, setSmtPartsSet] = useState<Set<string>>(new Set());
  const [readiness, setReadiness] = useState<ReadinessRow[] | null>(null);
  const [readinessLoading, setReadinessLoading] = useState(false);
  const [readinessError, setReadinessError] = useState<string | null>(null);
  const [expandedReadiness, setExpandedReadiness] = useState<Set<string>>(new Set());
  const [expandedPoRows, setExpandedPoRows] = useState<Set<string>>(new Set());
  const [showCompletedWos, setShowCompletedWos] = useState(false);
  const [bomTypeFilter, setBomTypeFilter] = useState<"ALL" | "PWB" | "HARNESS">("ALL");
  const [iopLookup, setIopLookup] = useState<Map<string, { nextDemand: string; schedule: string }>>(new Map());
  const [iopFileName, setIopFileName] = useState<string>("");
  const [matchedSort, setMatchedSort] = useState<{ col: string; asc: boolean }>({ col: "", asc: true });

  function toggleMatchedSort(col: string) {
    setMatchedSort((prev) => ({
      col,
      asc: prev.col === col ? !prev.asc : true,
    }));
  }

  // Load BOM database, floor stock, and WO status from Google Sheet
  useEffect(() => {
    void (async () => {
      const res = await fetch("/api/boms");
      const data = await res.json();
      setBoms(data.items || []);
    })();
    void (async () => {
      const res = await fetch("/api/import/floor-stock");
      const data = await res.json();
      setFloorStockSet(new Set((data.items || []).map((f: { component_pn: string }) => f.component_pn)));
    })();
    void (async () => {
      const res = await fetch("/api/import/open-po");
      const data = await res.json();
      setOpenPOs(data.items || []);
    })();
    void (async () => {
      try {
        const res = await fetch("/api/import/smt-parts");
        const data = await res.json();
        setSmtPartsSet(new Set((data.items || []).map((s: { component_pn: string }) => s.component_pn)));
      } catch { /* silent */ }
    })();
    // Helper: parse CSV rows with quoted field handling
    function parseCsvRows(text: string): string[][] {
      return text.split(/\r?\n/).filter(Boolean).map((line) => {
        const fields: string[] = [];
        let current = "";
        let inQuote = false;
        for (let i = 0; i < line.length; i++) {
          const ch = line[i];
          if (ch === '"') { inQuote = !inQuote; }
          else if (ch === ',' && !inQuote) { fields.push(current.trim()); current = ""; }
          else { current += ch; }
        }
        fields.push(current.trim());
        return fields;
      });
    }

    // Auto-fetch SMT WO status (Sheet 1 — existing sheet)
    void (async () => {
      try {
        const sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQrwgc_-77o7DUpMARlSl5mW72ZMJcEetuImpCDTbfFvdl4CRJODl9UrKx_NE8VSg/pub?output=csv";
        const res = await fetch(sheetUrl);
        if (!res.ok) return;
        const rows = parseCsvRows(await res.text());
        // Skip header rows (first two rows are headers)
        const dataRows = rows.slice(2);
        const map = new Map<string, string>();
        for (const r of dataRows) {
          const wo = (r[1] || "").trim(); // Column B = Work Order
          const status = (r[2] || "").trim(); // Column C = Status
          if (wo && status) map.set(wo, status);
        }
        if (map.size > 0) {
          setWoSmtStatusMap(map);
          setWoSmtStatusFileName(`Auto-loaded (${map.size} WOs)`);
        }
      } catch {
        // Silent fail
      }
    })();

    // Auto-fetch TH WO status (Sheet 2 — new through-hole sheet)
    void (async () => {
      try {
        const sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR4TXVWYps8hlXBXfu-Rwdel5BOKkIEQnYON7R0HUusTD7IvbWuLuMPonrabbR9jw/pub?output=csv";
        const res = await fetch(sheetUrl);
        if (!res.ok) return;
        const rows = parseCsvRows(await res.text());
        // First row is header: Board, Work order, Status, ...
        const dataRows = rows.slice(1);
        const map = new Map<string, string>();
        for (const r of dataRows) {
          const wo = (r[1] || "").trim(); // Column B = Work order
          const status = (r[2] || "").trim(); // Column C = Status
          if (wo && status) map.set(wo, status);
        }
        if (map.size > 0) {
          setWoThStatusMap(map);
          setWoThStatusFileName(`Auto-loaded (${map.size} WOs)`);
        }
      } catch {
        // Silent fail — user can still upload manually
      }
    })();
  }, []);

  // Build lookup: assembly_pn -> list of revisions (sorted, latest last)
  const filteredBoms = useMemo(() => {
    if (bomTypeFilter === "ALL") return boms;
    return boms.filter((b) => (b.bom_type || "PWB") === bomTypeFilter);
  }, [boms, bomTypeFilter]);

  const bomLookup = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const b of filteredBoms) {
      if (!map.has(b.assembly_pn)) map.set(b.assembly_pn, []);
      map.get(b.assembly_pn)!.push(b.revision);
    }
    for (const [, revs] of map) revs.sort();
    return map;
  }, [filteredBoms]);

  // A WO is fully complete if:
  //   - SMT sheet status is "Complete" (meaning SMT done, and if it's in the SMT-only sheet that tracks all phases), OR
  //   - TH sheet status is "completed" (meaning TH done = fully done)
  // A WO is "SMT complete" (partial) if SMT sheet = "Complete" but TH sheet != "completed"
  const isWoFullyComplete = useCallback((wo: string) => {
    const smtStatus = (woSmtStatusMap.get(wo) || "").trim().toLowerCase();
    const thStatus = (woThStatusMap.get(wo) || "").trim().toLowerCase();
    // Fully complete: SMT sheet says "complete" AND TH sheet says "completed"
    // OR TH sheet alone says "completed" (TH done = all done)
    return smtStatus === "complete" && thStatus === "completed";
  }, [woSmtStatusMap, woThStatusMap]);

  const isWoSmtComplete = useCallback((wo: string) => {
    const smtStatus = (woSmtStatusMap.get(wo) || "").trim().toLowerCase();
    return smtStatus === "complete";
  }, [woSmtStatusMap]);

  const getWoPhaseBadge = useCallback((woRaw: string): WoPhaseBadge => {
    const wo = woRaw.trim();
    if (!wo) {
      return {
        label: "MRP Plan",
        className: "text-xs px-1.5 py-0.5 rounded bg-amber-50 text-amber-800",
      };
    }

    const smtStatusRaw = (woSmtStatusMap.get(wo) || "").trim();
    const thStatusRaw = (woThStatusMap.get(wo) || "").trim();
    const smtStatus = smtStatusRaw.toLowerCase();
    const thStatus = thStatusRaw.toLowerCase();
    const smtDone = smtStatus === "complete";
    const thDone = thStatus === "completed";

    if (thDone && smtDone) {
      return {
        label: "✅ Done",
        className: "text-xs px-1.5 py-0.5 rounded bg-green-100 text-green-800",
      };
    }

    if (smtDone) {
      return {
        label: `SMT ✅ → TH ${thStatusRaw || "pending"}`,
        className: "text-xs px-1.5 py-0.5 rounded bg-blue-100 text-blue-800",
      };
    }

    if (smtStatusRaw) {
      return {
        label: smtStatusRaw,
        className: "text-xs px-1.5 py-0.5 rounded bg-amber-50 text-amber-800",
      };
    }

    if (thStatusRaw) {
      return {
        label: `TH ${thStatusRaw}`,
        className: "text-xs px-1.5 py-0.5 rounded bg-orange-100 text-orange-800",
      };
    }

    return {
      label: "SMT pending",
      className: "text-xs px-1.5 py-0.5 rounded bg-amber-50 text-amber-800",
    };
  }, [woSmtStatusMap, woThStatusMap]);

  const activeMatched = useMemo(() => {
    return matched.filter((m) => {
      const wo = (m.order || "").trim();
      return !isWoFullyComplete(wo);
    });
  }, [matched, isWoFullyComplete]);

  const completedMatched = useMemo(() => {
    return matched.filter((m) => {
      const wo = (m.order || "").trim();
      return isWoFullyComplete(wo);
    });
  }, [matched, isWoFullyComplete]);

  const activeMatchedIndexed = useMemo(
    () => matched.map((m, idx) => ({ m, idx })).filter(({ m }) => {
      const wo = (m.order || "").trim();
      return !isWoFullyComplete(wo);
    }),
    [matched, isWoFullyComplete]
  );

  const sortedMatchedIndexed = useMemo(() => {
    if (!matchedSort.col) return activeMatchedIndexed;
    const sorted = [...activeMatchedIndexed].sort((ea, eb) => {
      const a = ea.m;
      const b = eb.m;
      let va: string | number = "";
      let vb: string | number = "";
      switch (matchedSort.col) {
        case "part": va = a.part; vb = b.part; break;
        case "description": va = a.description; vb = b.description; break;
        case "order": va = a.order; vb = b.order; break;
        case "qty": va = a.qty; vb = b.qty; break;
        case "short": va = a.short; vb = b.short; break;
        case "iop": va = a.iop_next_demand || "9999"; vb = b.iop_next_demand || "9999"; break;
        case "priority": va = a.priority; vb = b.priority; break;
        default: return 0;
      }
      if (typeof va === "number" && typeof vb === "number") return matchedSort.asc ? va - vb : vb - va;
      return matchedSort.asc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
    return sorted;
  }, [activeMatchedIndexed, matchedSort]);

  const handleFiles = useCallback(
    (files: File[]) => {
      setParseError(null);
      setMpsRows([]);
      setMatched([]);
      setResult(null);
      setReadiness(null);
      setStockroomResults(null);
      setStockroomError(null);
      setExpandedPoRows(new Set());
      setExpandedReadiness(new Set());
      setCollapsedBoms(new Set());
      setSheetNames([]);
      setSelectedSheets(new Set());
      setWorkbooksRef([]);
      setFileNames(files.map((f) => f.name));

      const allWorkbooks: Array<{ name: string; wb: XLSX.WorkBook }> = [];
      let loaded = 0;

      for (const file of files) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const wb = XLSX.read(data, { type: "array" });
            allWorkbooks.push({ name: file.name, wb });
          } catch (err) {
            setParseError((prev) => (prev ? prev + "; " : "") + `${file.name}: ${err instanceof Error ? err.message : "Failed to read"}`);
          }
          loaded++;
          if (loaded === files.length) {
            // All files loaded — merge sheet names
            setWorkbooksRef(allWorkbooks);
            const allSheets: string[] = [];
            for (const { name, wb } of allWorkbooks) {
              for (const sn of wb.SheetNames) {
                const label = allWorkbooks.length > 1 ? `${sn} (${name})` : sn;
                allSheets.push(label);
              }
            }
            setSheetNames(allSheets);

            // Auto-parse if total 1 sheet across all files
            if (allSheets.length === 1) {
              const sel = new Set(allSheets);
              setSelectedSheets(sel);
              parseSheetsFromWorkbooks(allWorkbooks, sel);
            }
          }
        };
        reader.readAsArrayBuffer(file);
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [bomLookup]
  );

  function parseSheet(ws: XLSX.WorkSheet): MpsRow[] {
    const raw: (string | number | null)[][] = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: null,
    });

    // Find header row — look for "Lvl 0" or similar
    let headerIdx = -1;
    for (let i = 0; i < Math.min(raw.length, 10); i++) {
      const row = raw[i];
      if (row && row.some((c) => String(c || "").toLowerCase().includes("lvl"))) {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx === -1) return [];

    const headers = raw[headerIdx].map((c) => String(c || "").trim().toLowerCase());

    const levelCols: number[] = [];
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].match(/^lvl\s*\d+$/)) levelCols.push(i);
    }

    const descIdx = headers.findIndex((h) => h.includes("desc"));
    const qtyReqIdx = headers.findIndex((h) => h.includes("qty") && h.includes("req"));
    const availIdx = headers.findIndex((h) => h === "avail" || h.includes("avail"));
    const shortIdx = headers.findIndex((h) => h === "short" || h.includes("short"));
    const orderIdx = headers.findIndex((h) => h === "order" || h.includes("order"));
    const ordQtyIdx = headers.findIndex((h) => h.includes("ord") && h.includes("qty"));
    const dueIdx = headers.findIndex((h) => h.includes("due"));
    const remarksIdx = headers.findIndex((h) => h.includes("remark") || h.includes("sales"));

    const parsed: MpsRow[] = [];
    let lastPart = "";
    let lastLevel = -1;
    let lastDesc = "";
    let currentTopPn = "";
    for (let i = headerIdx + 1; i < raw.length; i++) {
      const row = raw[i];
      if (!row || row.every((c) => c === null || c === "")) continue;

      let level = -1;
      let part = "";
      for (const col of levelCols) {
        const val = String(row[col] || "").trim();
        if (val) {
          level = levelCols.indexOf(col);
          part = val;
          break;
        }
      }

      // If no part in any Lvl column, check if this is a sub-row (e.g. MRP Plan)
      // that should inherit the part from the previous row
      if (level === -1 || !part) {
        const order = orderIdx >= 0 ? String(row[orderIdx] || "").trim() : "";
        const ordQty = ordQtyIdx >= 0 ? Number(row[ordQtyIdx] || 0) : 0;
        const qtyReq = qtyReqIdx >= 0 ? Number(row[qtyReqIdx] || 0) : 0;
        // Only inherit if there's an order or qty value and we have a previous part
        if (lastPart && (order || ordQty > 0 || qtyReq > 0)) {
          part = lastPart;
          level = lastLevel;
        } else {
          continue;
        }
      } else {
        // Update last known part for inheritance
        lastPart = part;
        lastLevel = level;
        lastDesc = descIdx >= 0 ? String(row[descIdx] || "").trim() : "";
        // Track current top-level (Lvl 0) part number
        if (level === 0) {
          currentTopPn = part;
        }
      }

      const desc = descIdx >= 0 ? String(row[descIdx] || "").trim() : "";

      parsed.push({
        level,
        part,
        description: desc || lastDesc,
        qty_req: qtyReqIdx >= 0 ? Number(row[qtyReqIdx] || 0) : 0,
        avail: availIdx >= 0 ? Number(row[availIdx] || 0) : 0,
        short: shortIdx >= 0 ? Number(row[shortIdx] || 0) : 0,
        order: orderIdx >= 0 ? String(row[orderIdx] || "").trim() : "",
        ord_qty: ordQtyIdx >= 0 ? Number(row[ordQtyIdx] || 0) : 0,
        due_date: dueIdx >= 0 ? String(row[dueIdx] || "").trim() : "",
        remarks: remarksIdx >= 0 ? String(row[remarksIdx] || "").trim() : "",
        top_pn: currentTopPn,
      });
    }
    return parsed;
  }

  function parseSheetsFromWorkbooks(workbooks: Array<{ name: string; wb: XLSX.WorkBook }>, selectedLabels: Set<string>) {
    const allParsed: MpsRow[] = [];
    const failedSheets: string[] = [];

    for (const { name: fileName, wb } of workbooks) {
      for (const sheetName of wb.SheetNames) {
        const label = workbooks.length > 1 ? `${sheetName} (${fileName})` : sheetName;
        if (!selectedLabels.has(label)) continue;
        const ws = wb.Sheets[sheetName];
        if (!ws) continue;
        const rows = parseSheet(ws);
        if (rows.length === 0) {
          failedSheets.push(label);
        }
        allParsed.push(...rows);
      }
    }

    if (failedSheets.length > 0 && allParsed.length === 0) {
      setParseError(`Could not find header row (Lvl 0, Lvl 1, etc.) in: ${failedSheets.join(", ")}`);
    }

    setMpsRows(allParsed);

    // Cross-reference with BOM database — deduplicate same part + same WO
    const matches: MatchedBoard[] = [];
    const seen = new Set<string>();
    for (const mps of allParsed) {
      const revisions = bomLookup.get(mps.part);
      if (revisions && revisions.length > 0) {
        const dedupeKey = `${mps.part}||${mps.order}`;
        if (seen.has(dedupeKey)) continue;
        seen.add(dedupeKey);
        matches.push({
          part: mps.part,
          description: mps.description,
          qty: mps.ord_qty || mps.qty_req || 1,
          short: mps.short,
          order: mps.order,
          revision: revisions[revisions.length - 1],
          revisions,
          priority: 1,
          mps_level: mps.level,
          top_pn: mps.top_pn || "",
          iop_next_demand: iopLookup.get(mps.top_pn || "")?.nextDemand || "",
        });
      }
    }
    console.log("[MPS] matched boards:", matches.length, "unique top_pns:", [...new Set(matches.map(m => m.top_pn))], "with IOP:", matches.filter(m => m.iop_next_demand).length);
    setMatched(matches);
  }

  function parseWoStatusSheet(ws: XLSX.WorkSheet): Map<string, string> {
    const raw: (string | number | null)[][] = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: null,
    });

    const normalize = (v: string | number | null | undefined) => String(v || "").trim();
    const lower = (v: string | number | null | undefined) => normalize(v).toLowerCase();

    let headerIdx = -1;
    let workOrderCol = 1;
    let statusCol = 2;

    for (let i = 0; i < Math.min(raw.length, 30); i++) {
      const row = raw[i] || [];
      const lowers = row.map((c) => lower(c));
      const hasOrderHint = lowers.some((c) => c.includes("work order") || c === "order" || c.includes("order"));
      if (!hasOrderHint) continue;

      headerIdx = i;
      const detectedOrder = lowers.findIndex((c) => c.includes("work order") || c.includes("wo") || c === "order" || c.includes("order"));
      const detectedStatus = lowers.findIndex((c) => c.includes("status") || c.includes("state"));
      if (detectedOrder >= 0) workOrderCol = detectedOrder;
      if (detectedStatus >= 0) statusCol = detectedStatus;
      break;
    }

    const start = headerIdx >= 0 ? headerIdx + 1 : 0;
    const map = new Map<string, string>();
    for (let i = start; i < raw.length; i++) {
      const row = raw[i] || [];
      const wo = normalize(row[workOrderCol] ?? row[1]);
      const status = normalize(row[statusCol] ?? row[2]);
      if (!wo) continue;
      if (!status) continue;
      map.set(wo, status);
    }
    return map;
  }

  function handleIopFile(file: File) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        // Find the IOP sheet (name starts with "IOP")
        const iopSheetName = wb.SheetNames.find((s) => s.toUpperCase().startsWith("IOP"));
        if (!iopSheetName) { setParseError((prev) => (prev ? prev + "; " : "") + "No IOP sheet found"); return; }
        const ws = wb.Sheets[iopSheetName];
        // Use cellDates:true + raw:true so date headers come back as Date objects
        const raw: (string | number | Date | null | undefined)[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
        if (raw.length < 2) return;

        const header = raw[0];
        // Find date columns: headers that are Date objects or parseable date strings
        const dateCols: Array<{ idx: number; date: Date }> = [];
        for (let i = 0; i < (header?.length || 0); i++) {
          const h = header?.[i];
          if (h instanceof Date && !isNaN(h.getTime())) {
            dateCols.push({ idx: i, date: h });
          } else if (typeof h === "number" && h > 40000 && h < 60000) {
            // Excel serial date number — convert to JS Date
            const epoch = new Date(Date.UTC(1899, 11, 30));
            const d = new Date(epoch.getTime() + h * 86400000);
            if (!isNaN(d.getTime())) dateCols.push({ idx: i, date: d });
          } else if (typeof h === "string") {
            const d = new Date(h);
            if (!isNaN(d.getTime()) && d.getFullYear() >= 2020 && d.getFullYear() <= 2040) {
              dateCols.push({ idx: i, date: d });
            }
          }
        }
        dateCols.sort((a, b) => a.date.getTime() - b.date.getTime());
        console.log("[IOP] date columns found:", dateCols.length, dateCols.map(dc => dc.date.toISOString().slice(0,10)));

        const lookup = new Map<string, { nextDemand: string; schedule: string }>();
        for (let ri = 1; ri < raw.length; ri++) {
          const row = raw[ri];
          if (!row || !row[1]) continue;
          const itemRaw = String(row[1]).trim();
          // Extract canonical PN: everything before first " - " or first space
          const pn = itemRaw.split(/\s+-\s+/)[0].split(/\s+/)[0].trim();
          if (!pn || !pn.includes("-")) continue;

          let firstDemandDate: Date | null = null;
          const demandParts: string[] = [];
          for (const dc of dateCols) {
            const val = row[dc.idx];
            const qty = typeof val === "number" ? val : parseFloat(String(val || "0"));
            if (!isNaN(qty) && qty > 0) {
              if (!firstDemandDate) firstDemandDate = dc.date;
              const mm = String(dc.date.getMonth() + 1).padStart(2, "0");
              const dd = String(dc.date.getDate()).padStart(2, "0");
              demandParts.push(`${Math.round(qty)}×${mm}-${dd}`);
            }
          }

          if (firstDemandDate) {
            const dateStr = firstDemandDate.toISOString().slice(0, 10);
            // Only update if this is earlier than existing (same PN can appear multiple rows)
            const existing = lookup.get(pn);
            if (!existing || dateStr < existing.nextDemand) {
              lookup.set(pn, { nextDemand: dateStr, schedule: demandParts.join(", ") });
            }
          }
        }

        console.log("[IOP] parsed items:", lookup.size, [...lookup.entries()].slice(0, 10).map(([k,v]) => `${k} -> ${v.nextDemand}`));
        setIopLookup(lookup);
        setIopFileName(`${file.name} (${lookup.size} items)`);
      } catch (err) {
        setParseError((prev) => (prev ? prev + "; " : "") + `IOP (${file.name}): ${err instanceof Error ? err.message : "Failed to parse"}`);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleWoSmtStatusFile(file: File) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const merged = new Map<string, string>();
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const sheetMap = parseWoStatusSheet(ws);
          for (const [wo, status] of sheetMap.entries()) merged.set(wo, status);
        }
        setWoSmtStatusMap(merged);
        setWoSmtStatusFileName(file.name);
      } catch (err) {
        setParseError((prev) => (prev ? prev + "; " : "") + `SMT WO Status (${file.name}): ${err instanceof Error ? err.message : "Failed to parse"}`);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleWoThStatusFile(file: File) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const merged = new Map<string, string>();
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const sheetMap = parseWoStatusSheet(ws);
          for (const [wo, status] of sheetMap.entries()) merged.set(wo, status);
        }
        setWoThStatusMap(merged);
        setWoThStatusFileName(file.name);
      } catch (err) {
        setParseError((prev) => (prev ? prev + "; " : "") + `TH WO Status (${file.name}): ${err instanceof Error ? err.message : "Failed to parse"}`);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function toggleSheet(name: string) {
    setSelectedSheets((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name); else next.add(name);
      return next;
    });
  }

  function applySheetSelection() {
    if (workbooksRef.length === 0 || selectedSheets.size === 0) return;
    setResult(null);
    setReadiness(null);
    setStockroomResults(null);
    setStockroomError(null);
    setExpandedPoRows(new Set());
    setExpandedReadiness(new Set());
    setCollapsedBoms(new Set());
    parseSheetsFromWorkbooks(workbooksRef, selectedSheets);
  }

  function updateMatched(idx: number, patch: Partial<MatchedBoard>) {
    setMatched((prev) => prev.map((m, i) => (i === idx ? { ...m, ...patch } : m)));
  }

  function removeMatched(idx: number) {
    setMatched((prev) => prev.filter((_, i) => i !== idx));
  }

  useEffect(() => {
    setResult(null);
    setReadiness(null);
    setStockroomResults(null);
    setStockroomError(null);
    setExpandedPoRows(new Set());
    setExpandedReadiness(new Set());
    setCollapsedBoms(new Set());
  }, [matched]);

  // Patch IOP next-demand onto matched rows whenever the IOP lookup changes
  useEffect(() => {
    if (iopLookup.size === 0) return;
    console.log("[IOP] patching", iopLookup.size, "IOP items onto", matched.length, "matched rows");
    const topPns = new Set(matched.map(m => m.top_pn));
    console.log("[IOP] unique top_pn values in matched:", [...topPns]);
    console.log("[IOP] IOP keys:", [...iopLookup.keys()]);
    setMatched((prev) => prev.map((m) => {
      const iopEntry = iopLookup.get(m.top_pn);
      return { ...m, iop_next_demand: iopEntry?.nextDemand || "" };
    }));
  }, [iopLookup]);

  function buildAnalysisPlanRunId(m: MatchedBoard, idx: number) {
    return `${m.part}__${m.revision}__${m.order || "MRP"}__${idx}`;
  }

  function buildAnalysisPlanRunLabel(m: MatchedBoard) {
    const phase = m.order ? getWoPhaseBadge(m.order).label : "MRP Plan · SMT+TH";
    return m.order
      ? `${m.part} rev ${m.revision} · WO ${m.order} · ${phase}`
      : `${m.part} rev ${m.revision} · ${phase}`;
  }

  // Keep one analysis run per matched MPS line. Do not collapse same assembly/rev
  // across different WOs/phases, or status/shortage handling gets attributed to
  // the first matching row.
  function aggregatePlans() {
    const plans: Array<{ run_id: string; run_label: string; assembly_pn: string; revision: string; qty: number; priority: number }> = [];
    for (const { m, idx } of activeMatchedIndexed) {
      if (m.qty <= 0) continue;
      plans.push({
        run_id: buildAnalysisPlanRunId(m, idx),
        run_label: buildAnalysisPlanRunLabel(m),
        assembly_pn: m.part,
        revision: m.revision,
        qty: m.qty,
        priority: m.priority,
      });
    }
    return plans;
  }

  async function runAnalysis() {
    setLoading(true);
    setError(null);
    setResult(null);
    setStockroomResults(null);
    setStockroomError(null);
    try {
      const plans = aggregatePlans();

      if (plans.length === 0) throw new Error("No board assemblies to analyze");

      const res = await fetch("/api/analyze/multi-bom", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ plans }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.error || "Analysis failed");
      setResult(data);
    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : "Unknown error";
      setError(message);
    } finally {
      setLoading(false);
    }
  }

  async function checkBuildReadiness() {
    setReadinessLoading(true);
    setReadinessError(null);
    setReadiness(null);
    setExpandedReadiness(new Set());
    try {
      const plans = aggregatePlans();

      if (plans.length === 0) throw new Error("No board assemblies to check");

      // Step 1: Run multi-BOM analysis to get XRAY inventory shortages
      const analysisRes = await fetch("/api/analyze/multi-bom", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ plans }),
      });
      const analysisData = await analysisRes.json();
      if (!analysisRes.ok) throw new Error(analysisData?.error || "Analysis failed");

      // Step 2: Stockroom check — collect ALL shortage parts from individual plans
      const allShortageParts = new Map<string, { part: string; shortage: number; bom_run: string }>();
      for (const plan of analysisData.plans) {
        const bomRun = plan.run_label || `${plan.assembly_pn} rev ${plan.revision}`;
        for (const s of plan.shortages) {
          const key = `${s.part}||${bomRun}`;
          if (!allShortageParts.has(key)) {
            allShortageParts.set(key, { part: s.part, shortage: s.shortage, bom_run: bomRun });
          }
        }
      }
      const stockroomMap = new Map<string, number>();

      if (allShortageParts.size > 0) {
        const srRes = await fetch("/api/analyze/stockroom-check", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            shortages: Array.from(allShortageParts.values()),
          }),
        });
        const srData = await srRes.json();
        if (srRes.ok && srData.results) {
          for (const r of srData.results) {
            stockroomMap.set(`${r.part}||${r.bom_run}`, r.stockroom_available);
            stockroomMap.set(r.part, r.stockroom_available);
          }
        }
      }

      // Step 3: Build readiness per assembly
      const rows: ReadinessRow[] = [];
      for (const plan of analysisData.plans) {
        const matchedEntry = activeMatchedIndexed.find(({ m, idx }) => buildAnalysisPlanRunId(m, idx) === plan.run_id)?.m
          || activeMatched.find((m) => m.part === plan.assembly_pn);
        const description = matchedEntry?.description || "";
        const woNumber = (matchedEntry?.order || "").trim();
        const woIsSmtComplete = isWoSmtComplete(woNumber);

        const details: ReadinessRow["details"] = [];
        let hasShortageAfterStockroom = false;

        // For each shortage part in this assembly
        for (const s of plan.shortages) {
          const isSMT = smtPartsSet.has(s.part);
          // If WO is SMT-complete, treat SMT parts like floor stock (no shortage)
          const smtFlushed = woIsSmtComplete && isSMT;
          const isFI = floorStockSet.has(s.part) || smtFlushed;
          const srKey = `${s.part}||${plan.run_label || `${plan.assembly_pn} rev ${plan.revision}`}`;
          const srAvail = stockroomMap.get(srKey) ?? stockroomMap.get(s.part) ?? 0;
          const combinedAvail = s.available + srAvail;
          const remainingShortage = isFI ? 0 : Math.max(0, s.required - combinedAvail);

          if (remainingShortage > 0) hasShortageAfterStockroom = true;

          details.push({
            part: s.part,
            required: s.required,
            xray_available: s.available,
            stockroom_available: srAvail,
            combined_available: combinedAvail,
            shortage: remainingShortage,
            is_floor_stock: isFI,
            mount_type: smtPartsSet.has(s.part) ? "SMT" : "TH",
            is_shared_conflict: false,
            is_shared: false,
          });
        }

        // Calculate max_buildable_combined (skip FI parts — always available)
        let maxBuildableCombined = plan.max_buildable;
        const nonFiDetails = details.filter((d) => !d.is_floor_stock);
        if (nonFiDetails.length > 0 && plan.qty > 0) {
          let minBoardsCombined = Number.MAX_SAFE_INTEGER;
          for (const d of nonFiDetails) {
            const qtyPerBoard = d.required / plan.qty;
            if (qtyPerBoard > 0) {
              const boards = Math.floor(d.combined_available / qtyPerBoard);
              minBoardsCombined = Math.min(minBoardsCombined, boards);
            }
          }
          if (minBoardsCombined !== Number.MAX_SAFE_INTEGER) {
            maxBuildableCombined = Math.min(
              plan.qty,
              nonFiDetails.length > 0 ? Math.max(plan.max_buildable, minBoardsCombined) : plan.max_buildable
            );
          }
        } else if (nonFiDetails.length === 0 && details.length > 0) {
          // All shortage parts are FI — fully buildable
          maxBuildableCombined = plan.qty;
        }

        const nonFiShortageCount = details.filter((d) => d.shortage > 0 && !d.is_floor_stock).length;
        const allShortsAreFI = plan.shortages.length > 0 && nonFiShortageCount === 0 && !hasShortageAfterStockroom;

        let status: ReadinessStatus;
        if (plan.can_complete || (plan.shortages.length > 0 && details.every((d) => d.is_floor_stock))) {
          status = "clear-xray";
        } else if (!hasShortageAfterStockroom || allShortsAreFI) {
          status = "clear-with-stockroom";
        } else {
          status = "not-clear";
        }

        rows.push({
          run_id: plan.run_id,
          run_label: plan.run_label,
          assembly_pn: plan.assembly_pn,
          revision: plan.revision,
          description,
          qty: plan.qty,
          status,
          max_buildable_xray: plan.max_buildable,
          max_buildable_combined: maxBuildableCombined,
          shortage_parts_xray: plan.shortages.length,
          shortage_parts_after_stockroom: nonFiShortageCount,
          details: details.sort((a, b) => b.shortage - a.shortage),
        });
      }

      // Detect shared and shared-short parts across assemblies
      if (rows.length > 1) {
        const totalDemand = new Map<string, number>();
        const totalSupply = new Map<string, number>();
        for (const row of rows) {
          for (const d of row.details) {
            if (d.is_floor_stock) continue;
            totalDemand.set(d.part, (totalDemand.get(d.part) || 0) + d.required);
            if (!totalSupply.has(d.part)) {
              totalSupply.set(d.part, d.combined_available);
            }
          }
        }
        const sharedParts = new Set<string>();
        const conflictParts = new Set<string>();
        for (const [part, demand] of totalDemand) {
          let count = 0;
          for (const row of rows) {
            if (row.details.some((d) => d.part === part && !d.is_floor_stock)) count++;
          }
          if (count > 1) {
            sharedParts.add(part);
            const supply = totalSupply.get(part) || 0;
            if (demand > supply) conflictParts.add(part);
          }
        }
        for (const row of rows) {
          for (const d of row.details) {
            if (conflictParts.has(d.part)) { d.is_shared_conflict = true; d.is_shared = true; }
            else if (sharedParts.has(d.part)) { d.is_shared = true; }
          }
        }
      }

      // Sort: clear-xray first, then clear-with-stockroom, then not-clear
      const statusOrder: Record<ReadinessStatus, number> = { "clear-xray": 0, "clear-with-stockroom": 1, "not-clear": 2 };
      rows.sort((a, b) => statusOrder[a.status] - statusOrder[b.status]);

      setReadiness(rows);
    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : "Unknown error";
      setReadinessError(message);
    } finally {
      setReadinessLoading(false);
    }
  }

  // Detect parts shared across multiple BOM runs and whether total demand exceeds supply
  const sharedPartsInfo = useMemo(() => {
    const map = new Map<string, { bomRuns: Set<string>; totalShortage: number; stockroomAvail: number }>();
    if (!stockroomResults) return map;
    for (const r of stockroomResults) {
      if (floorStockSet.has(r.part)) continue;
      if (!map.has(r.part)) map.set(r.part, { bomRuns: new Set(), totalShortage: 0, stockroomAvail: r.stockroom_available });
      const entry = map.get(r.part)!;
      entry.bomRuns.add(r.bom_run);
      entry.totalShortage += r.shortage_qty;
    }
    // Only keep parts that appear in 2+ BOM runs
    for (const [part, info] of map) {
      if (info.bomRuns.size < 2) map.delete(part);
    }
    return map;
  }, [stockroomResults, floorStockSet]);

  const sortedStockroomResults = useMemo(() => {
    if (!stockroomResults) return [];
    const sorted = [...stockroomResults];
    const { col, asc } = stockroomSort;
    sorted.sort((a, b) => {
      let va: string | number, vb: string | number;
      const statusRank = (r: StockroomResult) => {
        if (floorStockSet.has(r.part)) return 4;
        const conf = r.confirmed_qty ?? r.stockroom_available;
        if (conf >= r.shortage_qty) return 3;
        if (conf + (r.inspection_qty ?? 0) >= r.shortage_qty) return 2;
        if (conf > 0 || (r.inspection_qty ?? 0) > 0) return 1;
        return 0;
      };
      switch (col) {
        case "part": va = a.part; vb = b.part; break;
        case "bom_run": va = a.bom_run; vb = b.bom_run; break;
        case "shortage_qty": va = a.shortage_qty; vb = b.shortage_qty; break;
        case "confirmed_qty": va = a.confirmed_qty ?? a.stockroom_available; vb = b.confirmed_qty ?? b.stockroom_available; break;
        case "mount_type": va = a.mount_type || "TH"; vb = b.mount_type || "TH"; break;
        case "inspection_qty": va = a.inspection_qty ?? 0; vb = b.inspection_qty ?? 0; break;
        case "shared": va = sharedPartsInfo.has(a.part) ? 1 : 0; vb = sharedPartsInfo.has(b.part) ? 1 : 0; break;
        case "status": va = statusRank(a); vb = statusRank(b); break;
        default: va = a.shortage_qty; vb = b.shortage_qty;
      }
      if (typeof va === "string" && typeof vb === "string") return asc ? va.localeCompare(vb) : vb.localeCompare(va);
      return asc ? (va as number) - (vb as number) : (vb as number) - (va as number);
    });
    return sorted;
  }, [stockroomResults, stockroomSort, sharedPartsInfo, floorStockSet]);

  function toggleStockroomSort(col: string) {
    setStockroomSort((prev) => prev.col === col ? { col, asc: !prev.asc } : { col, asc: col === "part" || col === "bom_run" });
  }

  function sortIndicator(col: string) {
    if (stockroomSort.col !== col) return " ↕";
    return stockroomSort.asc ? " ↑" : " ↓";
  }

  function toggleReadinessExpand(key: string) {
    setExpandedReadiness((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  }

  // --- Shortage helpers (same as multi-bom page) ---

  const shortagesByBom = useMemo(() => {
    if (!result) return [];
    const groups = new Map<string, { bom_run: string; priority: number; shortages: typeof result.top_shortages }>();
    for (const s of result.top_shortages) {
      if (!groups.has(s.bom_run)) groups.set(s.bom_run, { bom_run: s.bom_run, priority: s.priority, shortages: [] });
      groups.get(s.bom_run)!.shortages.push(s);
    }
    for (const g of groups.values()) g.shortages.sort((a, b) => b.shortage - a.shortage);
    return Array.from(groups.values()).sort((a, b) => b.priority - a.priority);
  }, [result]);

  const crossBomShortages = useMemo(() => {
    if (!result) return [];
    const partMap = new Map<string, { part: string; totalShortage: number; bomRuns: string[]; maxPriority: number }>();
    for (const s of result.top_shortages) {
      if (!partMap.has(s.part)) partMap.set(s.part, { part: s.part, totalShortage: 0, bomRuns: [], maxPriority: 0 });
      const entry = partMap.get(s.part)!;
      entry.totalShortage += s.shortage;
      if (!entry.bomRuns.includes(s.bom_run)) entry.bomRuns.push(s.bom_run);
      entry.maxPriority = Math.max(entry.maxPriority, s.priority);
    }
    return Array.from(partMap.values()).sort((a, b) => b.totalShortage - a.totalShortage);
  }, [result]);

  const openPOLookup = useMemo(() => {
    const map = new Map<string, typeof openPOs>();
    for (const po of openPOs) {
      const key = po.item.toUpperCase();
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(po);
    }
    return map;
  }, [openPOs]);


  function toggleBomCollapse(bomKey: string) {
    setCollapsedBoms((prev) => {
      const next = new Set(prev);
      if (next.has(bomKey)) next.delete(bomKey); else next.add(bomKey);
      return next;
    });
  }

  function togglePoExpand(key: string) {
    setExpandedPoRows((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  }

  async function checkStockroom() {
    if (!result) return;
    setStockroomLoading(true);
    setStockroomError(null);
    setStockroomResults(null);
    try {
      // Collect shortage parts from individual plans so each WO/MRP run stays separate.
      const seen = new Set<string>();
      const shortages: Array<{ part: string; shortage: number; bom_run: string }> = [];
      for (const plan of result.plans) {
        const bomRun = plan.run_label || `${plan.assembly_pn} rev ${plan.revision}`;
        // Check if this WO is SMT-complete
        const matchedEntry = activeMatchedIndexed.find(({ m, idx }) => buildAnalysisPlanRunId(m, idx) === plan.run_id)?.m
          || activeMatched.find((m) => m.part === plan.assembly_pn);
        const woNumber = (matchedEntry?.order || "").trim();
        const woSmtDone = isWoSmtComplete(woNumber);
        for (const s of plan.shortages) {
          // Skip SMT parts if WO is SMT-complete (already consumed)
          if (woSmtDone && smtPartsSet.has(s.part)) continue;
          // Skip floor stock parts
          if (floorStockSet.has(s.part)) continue;
          const key = `${s.part}||${bomRun}`;
          if (!seen.has(key)) {
            seen.add(key);
            shortages.push({ part: s.part, shortage: s.shortage, bom_run: bomRun });
          }
        }
      }
      const res = await fetch("/api/analyze/stockroom-check", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ shortages }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.error || "Stockroom check failed");
      const srResults: StockroomResult[] = (data.results || []).map((r: StockroomResult) => ({
        ...r,
        mount_type: smtPartsSet.has(r.part) ? "SMT" as const : "TH" as const,
      }));
      setStockroomResults(srResults);
    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : "Unknown error";
      setStockroomError(message);
    } finally {
      setStockroomLoading(false);
    }
  }

  function downloadCsv(filename: string, csvContent: string) {
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  function exportCsv() {
    if (!result) return;
    const header = ["run", "assembly_pn", "revision", "qty", "priority", "can_complete", "max_buildable", "recommended"];
    const lines = result.plans.map((p) => [p.run_label || `${p.assembly_pn} rev ${p.revision}`, p.assembly_pn, p.revision, p.qty, p.priority, p.can_complete ? "YES" : "NO", p.max_buildable, p.in_recommended ? "YES" : "NO"]);
    downloadCsv("mps-analysis.csv", [header, ...lines].map((r) => r.join(",")).join("\n"));
  }

  function exportShortagesCsv() {
    if (!result) return;
    const header = ["part", "shortage", "required", "available", "bom_run", "priority"];
    const lines = result.top_shortages.map((s) => [s.part, s.shortage, s.required, s.available, s.bom_run, s.priority]);
    downloadCsv("mps-shortage-parts.csv", [header, ...lines].map((r) => r.join(",")).join("\n"));
  }

  function exportStockroomCsv() {
    if (!stockroomResults) return;
    const header = ["status", "mount_type", "part", "bom_run", "shortage", "confirmed_stockroom_qty", "in_inspection", "shared_boms", "contested", "on_order_qty", "vendor", "po", "line", "promise_date", "original_promise", "description", "locations"];
    const lines = stockroomResults.map((row) => {
      const isFI = floorStockSet.has(row.part);
      const shared = sharedPartsInfo.get(row.part);
      const conf = row.confirmed_qty ?? row.stockroom_available;
      const insp = row.inspection_qty ?? 0;
      const poList = openPOLookup.get(row.part.toUpperCase()) || [];
      const poTotal = poList.reduce((sum, po) => sum + po.qty_due, 0);
      const vendors = [...new Set(poList.map((po) => po.vendor).filter(Boolean))].join("; ");
      const pos = poList.map((po) => po.po).filter(Boolean).join("; ");
      const poLines = poList.map((po) => po.line).filter(Boolean).join("; ");
      const promises = poList.map((po) => po.promise_date).filter(Boolean).join("; ");
      const origPromises = poList.map((po) => po.original_promise).filter(Boolean).join("; ");
      const descriptions = [...new Set(poList.map((po) => po.description).filter(Boolean))].join("; ");
      const locations = isFI
        ? "Floor Stock (FI)"
        : row.stockroom_locations.map((l) => `${l.qty} @ ${l.bin_location} (${l.lot_number})`).join("; ");
      const status = isFI ? "FI" : conf >= row.shortage_qty ? "Covered" : (conf + insp) >= row.shortage_qty ? "Covered (needs inspection)" : conf > 0 || insp > 0 ? "Partial" : "Not in stockroom";
      return [
        csvCell(status),
        csvCell(row.mount_type || "TH"),
        csvCell(row.part),
        csvCell(row.bom_run),
        csvCell(isFI ? "" : row.shortage_qty),
        csvCell(isFI ? "" : conf),
        csvCell(isFI ? "" : insp > 0 ? insp : ""),
        csvCell(shared ? shared.bomRuns.size : ""),
        csvCell(shared && shared.totalShortage > shared.stockroomAvail ? "YES" : ""),
        csvCell(!isFI && poList.length > 0 ? poTotal : ""),
        csvCell(vendors),
        csvCell(pos),
        csvCell(poLines),
        csvCell(promises),
        csvCell(origPromises),
        csvCell(descriptions),
        csvCell(locations),
      ];
    });
    downloadCsv("mps-stockroom-availability.csv", [header.map(csvCell), ...lines].map((r) => r.join(",")).join("\n"));
  }

  function exportStockroomShortagesOnlyCsv() {
    if (!stockroomResults) return;
    const shortOnly = stockroomResults.filter((row) => {
      const isFI = floorStockSet.has(row.part);
      const conf = row.confirmed_qty ?? row.stockroom_available;
      return !isFI && conf < row.shortage_qty;
    });
    const header = ["status", "mount_type", "part", "bom_run", "shortage", "confirmed_stockroom_qty", "in_inspection", "shared_boms", "contested", "on_order_qty", "vendor", "po", "line", "promise_date", "original_promise", "description", "locations"];
    const lines = shortOnly.map((row) => {
      const shared = sharedPartsInfo.get(row.part);
      const conf = row.confirmed_qty ?? row.stockroom_available;
      const insp = row.inspection_qty ?? 0;
      const poList = openPOLookup.get(row.part.toUpperCase()) || [];
      const poTotal = poList.reduce((sum, po) => sum + po.qty_due, 0);
      const vendors = [...new Set(poList.map((po) => po.vendor).filter(Boolean))].join("; ");
      const pos = poList.map((po) => po.po).filter(Boolean).join("; ");
      const poLines = poList.map((po) => po.line).filter(Boolean).join("; ");
      const promises = poList.map((po) => po.promise_date).filter(Boolean).join("; ");
      const origPromises = poList.map((po) => po.original_promise).filter(Boolean).join("; ");
      const descriptions = [...new Set(poList.map((po) => po.description).filter(Boolean))].join("; ");
      const locations = row.stockroom_locations.map((l) => `${l.qty} @ ${l.bin_location} (${l.lot_number})`).join("; ");
      const status = conf >= row.shortage_qty ? "Covered" : (conf + insp) >= row.shortage_qty ? "Covered (needs inspection)" : conf > 0 || insp > 0 ? "Partial" : "Not in stockroom";
      return [
        csvCell(status),
        csvCell(row.mount_type || "TH"),
        csvCell(row.part),
        csvCell(row.bom_run),
        csvCell(row.shortage_qty),
        csvCell(conf),
        csvCell(insp > 0 ? insp : ""),
        csvCell(shared ? shared.bomRuns.size : ""),
        csvCell(shared && shared.totalShortage > shared.stockroomAvail ? "YES" : ""),
        csvCell(poList.length > 0 ? poTotal : ""),
        csvCell(vendors),
        csvCell(pos),
        csvCell(poLines),
        csvCell(promises),
        csvCell(origPromises),
        csvCell(descriptions),
        csvCell(locations),
      ];
    });
    downloadCsv("mps-stockroom-shortages-only.csv", [header.map(csvCell), ...lines].map((r) => r.join(",")).join("\n"));
  }

  // Drag & drop handlers
  function handleDrop(e: React.DragEvent) {
    e.preventDefault();
    const files = Array.from(e.dataTransfer.files).filter((f) => f.name.match(/\.(xlsx|xls|csv)$/i));
    if (files.length > 0) handleFiles(files);
  }

  return (
    <main className="max-w-5xl mx-auto p-6 space-y-6 w-full">
      <h1 className="text-2xl font-bold">MPS Import</h1>
      <p className="text-sm text-gray-600">
        Upload an MPS Excel file. Board assemblies matching your BOM database are auto-detected with quantities pre-filled. Assign priorities, then run the Multi-BOM analysis.
      </p>

      <div className="flex items-center gap-2 text-sm">
        <span className="text-gray-600">BOM Filter:</span>
        {(["ALL", "PWB", "HARNESS"] as const).map((t) => (
          <button
            key={t}
            onClick={() => setBomTypeFilter(t)}
            className={`px-3 py-1 rounded text-xs font-medium transition-colors ${
              bomTypeFilter === t
                ? t === "PWB" ? "bg-blue-600 text-white" : t === "HARNESS" ? "bg-amber-600 text-white" : "bg-black text-white"
                : "bg-gray-100 hover:bg-gray-200 text-gray-700"
            }`}
          >
            {t === "ALL" ? `All (${boms.length})` : t === "PWB" ? `PWB (${boms.filter(b => (b.bom_type||"PWB")==="PWB").length})` : `Harness (${boms.filter(b => b.bom_type==="HARNESS").length})`}
          </button>
        ))}
      </div>

      {/* Upload section */}
      <section
        className="border-2 border-dashed rounded-lg p-8 text-center cursor-pointer hover:border-gray-400 transition-colors"
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
        onClick={() => document.getElementById("mps-file-input")?.click()}
      >
        <input
          id="mps-file-input"
          type="file"
          accept=".xlsx,.xls,.csv"
          multiple
          className="hidden"
          onChange={(e) => {
            const files = Array.from(e.target.files || []);
            if (files.length > 0) handleFiles(files);
          }}
        />
        {fileNames.length > 0 ? (
          <div>
            <p className="text-lg font-medium">📄 {fileNames.length === 1 ? fileNames[0] : `${fileNames.length} files`}</p>
            {fileNames.length > 1 && (
              <p className="text-sm text-gray-500 mt-1">{fileNames.join(", ")}</p>
            )}
            <p className="text-sm text-gray-500 mt-1">
              {mpsRows.length} rows parsed · {activeMatched.length} active board assembl{activeMatched.length === 1 ? "y" : "ies"} matched
              {completedMatched.length > 0 && ` · ${completedMatched.length} completed filtered`}
            </p>
            <p className="text-xs text-gray-400 mt-2">Drop more files to replace</p>
          </div>
        ) : (
          <div>
            <p className="text-lg text-gray-500">Drop MPS Excel file(s) here or click to browse</p>
            <p className="text-xs text-gray-400 mt-2">Supports .xlsx, .xls, .csv — multiple files OK</p>
          </div>
        )}
      </section>

      {parseError && <div className="text-sm text-red-600 border border-red-200 rounded p-3">{parseError}</div>}

      <section className="border rounded p-3 space-y-2">
        <h2 className="font-semibold">WO Phase Status</h2>
        <p className="text-xs text-gray-500">Auto-loaded from Google Sheets. SMT Complete → flushes SMT parts. TH Complete → WO fully done. Upload to override.</p>
        <div className="grid grid-cols-2 gap-4">
          <div className="border rounded p-2 space-y-1">
            <div className="text-sm font-medium text-blue-800">🟦 SMT Status</div>
            <div className="flex items-center gap-2">
              {woSmtStatusFileName ? (
                <span className="text-xs">✅ {woSmtStatusFileName} ({woSmtStatusMap.size} WOs)</span>
              ) : (
                <span className="text-xs text-gray-400">Loading...</span>
              )}
              <div
                className="border border-dashed rounded px-2 py-0.5 text-xs cursor-pointer hover:border-blue-400 transition-colors"
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleWoSmtStatusFile(f); }}
                onClick={() => document.getElementById("wo-smt-file")?.click()}
              >
                <input id="wo-smt-file" type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) handleWoSmtStatusFile(f); }} />
                Override
              </div>
            </div>
          </div>
          <div className="border rounded p-2 space-y-1">
            <div className="text-sm font-medium text-gray-700">⬛ TH Status</div>
            <div className="flex items-center gap-2">
              {woThStatusFileName ? (
                <span className="text-xs">✅ {woThStatusFileName} ({woThStatusMap.size} WOs)</span>
              ) : (
                <span className="text-xs text-gray-400">Loading...</span>
              )}
              <div
                className="border border-dashed rounded px-2 py-0.5 text-xs cursor-pointer hover:border-gray-400 transition-colors"
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleWoThStatusFile(f); }}
                onClick={() => document.getElementById("wo-th-file")?.click()}
              >
                <input id="wo-th-file" type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) handleWoThStatusFile(f); }} />
                Override
              </div>
            </div>
          </div>
        </div>
      </section>

      <section className="border rounded p-3 space-y-2">
        <h2 className="font-semibold">IOP Demand Schedule</h2>
        <p className="text-xs text-gray-500">Upload the IOP Excel file to map demand dates onto matched board assemblies via their top-level unit part number.</p>
        <div className="flex items-center gap-2">
          {iopFileName ? (
            <span className="text-xs">✅ {iopFileName}</span>
          ) : (
            <span className="text-xs text-gray-400">No IOP file loaded</span>
          )}
          <div
            className="border border-dashed rounded px-3 py-1 text-xs cursor-pointer hover:border-blue-400 transition-colors"
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleIopFile(f); }}
            onClick={() => document.getElementById("iop-file-input")?.click()}
          >
            <input id="iop-file-input" type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) handleIopFile(f); }} />
            {iopFileName ? "Replace" : "Upload IOP"}
          </div>
        </div>
      </section>

      {openPOs.length > 0 && (
        <section className="border rounded p-3 space-y-2">
          <h2 className="font-semibold">Open Purchase Orders</h2>
          <p className="text-xs text-gray-500">Loaded from Data Manager, so both MPS Import and Multi-BOM use the same PO source.</p>
          <div className="text-sm">✅ {openPOs.length} open PO lines loaded</div>
        </section>
      )}

      {/* Sheet selector — shown when multiple sheets exist */}
      {sheetNames.length > 1 && (
        <section className="border rounded p-3 space-y-3">
          <h2 className="font-semibold">Select Sheets to Analyze</h2>
          <p className="text-xs text-gray-500">This workbook has {sheetNames.length} tabs. Select one or more to include in the analysis.</p>
          <div className="flex flex-wrap gap-2">
            {sheetNames.map((name) => {
              const isSelected = selectedSheets.has(name);
              return (
                <button
                  key={name}
                  className={`px-3 py-2 rounded border text-sm font-medium transition-colors ${
                    isSelected
                      ? "bg-black text-white border-black"
                      : "bg-white text-gray-700 border-gray-300 hover:border-gray-500"
                  }`}
                  onClick={() => toggleSheet(name)}
                >
                  {name}
                </button>
              );
            })}
          </div>
          <div className="flex items-center gap-3">
            <button
              className="rounded bg-black text-white px-4 py-2 text-sm disabled:opacity-50"
              onClick={applySheetSelection}
              disabled={selectedSheets.size === 0}
            >
              Load Selected ({selectedSheets.size})
            </button>
            <button
              className="rounded border px-3 py-1 text-sm"
              onClick={() => {
                setSelectedSheets(new Set(sheetNames));
              }}
            >
              Select All
            </button>
            <button
              className="rounded border px-3 py-1 text-sm"
              onClick={() => setSelectedSheets(new Set())}
            >
              Clear
            </button>
          </div>
        </section>
      )}

      {/* MPS overview — all rows parsed */}
      {mpsRows.length > 0 && (
        <section className="border rounded p-3 space-y-2">
          <h2 className="font-semibold">MPS Overview ({mpsRows.length} rows)</h2>
          <div className="overflow-auto max-h-64">
            <table className="w-full text-xs">
              <thead>
                <tr className="text-left border-b">
                  <th className="p-1">Lvl</th>
                  <th className="p-1">Part Number</th>
                  <th className="p-1">Description</th>
                  <th className="p-1">Qty Req</th>
                  <th className="p-1">Short</th>
                  <th className="p-1">In BOM DB</th>
                </tr>
              </thead>
              <tbody>
                {mpsRows.map((r, i) => {
                  const inBom = bomLookup.has(r.part);
                  return (
                    <tr key={i} className={`border-b ${inBom ? "bg-green-50" : ""}`}>
                      <td className="p-1">{r.level}</td>
                      <td className="p-1 font-mono">{r.part}</td>
                      <td className="p-1">{r.description}</td>
                      <td className="p-1">{r.qty_req}</td>
                      <td className="p-1">{r.short}</td>
                      <td className="p-1">{inBom ? "✅" : "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {/* Matched board assemblies */}
      {activeMatched.length > 0 && (
        <section className="border rounded p-3 space-y-3">
          <h2 className="font-semibold">Matched Board Assemblies ({activeMatched.length})</h2>
          <p className="text-xs text-gray-500">These MPS parts match your BOM database. Qty is pre-filled from MPS. Assign priority (higher = first) then run analysis.</p>

          <div className="overflow-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left border-b">
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-50" onClick={() => toggleMatchedSort("part")}>
                    Part Number {matchedSort.col === "part" ? (matchedSort.asc ? "▲" : "▼") : ""}
                  </th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-50" onClick={() => toggleMatchedSort("description")}>
                    Description {matchedSort.col === "description" ? (matchedSort.asc ? "▲" : "▼") : ""}
                  </th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-50" onClick={() => toggleMatchedSort("order")}>
                    Work Order {matchedSort.col === "order" ? (matchedSort.asc ? "▲" : "▼") : ""}
                  </th>
                  <th className="p-2">Phase</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-50" onClick={() => toggleMatchedSort("iop")}>
                    IOP Demand {matchedSort.col === "iop" ? (matchedSort.asc ? "▲" : "▼") : ""}
                  </th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-50" onClick={() => toggleMatchedSort("qty")}>
                    Ord Qty {matchedSort.col === "qty" ? (matchedSort.asc ? "▲" : "▼") : ""}
                  </th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-50" onClick={() => toggleMatchedSort("short")}>
                    MPS Short {matchedSort.col === "short" ? (matchedSort.asc ? "▲" : "▼") : ""}
                  </th>
                  <th className="p-2">Revision</th>
                  <th className="p-2">Target Qty</th>
                  <th className="p-2 cursor-pointer select-none hover:bg-gray-50" onClick={() => toggleMatchedSort("priority")}>
                    Priority {matchedSort.col === "priority" ? (matchedSort.asc ? "▲" : "▼") : ""}
                  </th>
                  <th className="p-2">Action</th>
                </tr>
              </thead>
              <tbody>
                {(() => {
                  // Group matched rows by part number
                  const groups: Array<{ part: string; indices: number[] }> = [];
                  const partGroupMap = new Map<string, number>();
                  sortedMatchedIndexed.forEach(({ m }, i) => {
                    const existing = partGroupMap.get(m.part);
                    if (existing !== undefined) {
                      groups[existing].indices.push(i);
                    } else {
                      partGroupMap.set(m.part, groups.length);
                      groups.push({ part: m.part, indices: [i] });
                    }
                  });

                  return groups.flatMap((group) => {
                    const isGrouped = group.indices.length > 1;
                    const firstIdx = group.indices[0];
                    const firstItem = sortedMatchedIndexed[firstIdx];
                    const first = firstItem.m;
                    const totalQty = group.indices.reduce((sum, idx) => sum + sortedMatchedIndexed[idx].m.qty, 0);

                    if (!isGrouped) {
                      // Single row — render normally
                      const m = first;
                      const i = firstItem.idx;
                      return [
                        <tr key={`${m.part}-${m.order}-${i}`} className="border-b">
                          <td className="p-2 font-mono">{m.part}</td>
                          <td className="p-2 text-xs">{m.description}</td>
                          <td className="p-2 text-xs">{m.order || "—"}</td>
                          <td className="p-2">
                            {(() => {
                              const badge = getWoPhaseBadge(m.order || "");
                              return <span className={badge.className}>{badge.label}</span>;
                            })()}
                          </td>
                          <td className="p-2 text-xs">
                            {m.iop_next_demand ? (
                              <span className="text-xs px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-800">
                                {new Date(m.iop_next_demand + "T00:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })}
                              </span>
                            ) : (
                              <span className="text-gray-300">—</span>
                            )}
                          </td>
                          <td className="p-2 text-gray-500">{m.qty}</td>
                          <td className="p-2 text-gray-500">{m.short}</td>
                          <td className="p-2">
                            {m.revisions.length > 1 ? (
                              <select
                                className="border rounded p-1 text-sm"
                                value={m.revision}
                                onChange={(e) => updateMatched(i, { revision: e.target.value })}
                              >
                                {m.revisions.map((r) => (
                                  <option key={r} value={r}>Rev {r}</option>
                                ))}
                              </select>
                            ) : (
                              <span>Rev {m.revision}</span>
                            )}
                          </td>
                          <td className="p-2">
                            <input
                              className="border rounded p-1 w-20 text-sm"
                              type="number"
                              min={1}
                              value={m.qty}
                              onChange={(e) => updateMatched(i, { qty: Number(e.target.value || 1) })}
                            />
                          </td>
                          <td className="p-2">
                            <input
                              className="border rounded p-1 w-20 text-sm"
                              type="number"
                              min={0}
                              step={1}
                              value={m.priority}
                              onChange={(e) => updateMatched(i, { priority: Number(e.target.value || 1) })}
                            />
                          </td>
                          <td className="p-2">
                            <button className="rounded border px-2 py-1 text-xs" onClick={() => removeMatched(i)}>Remove</button>
                          </td>
                        </tr>,
                      ];
                    }

                    // Grouped rows — part number header + sub-rows per WO
                    return [
                      // Group header row
                      <tr key={`group-header-${group.part}`} className="border-b bg-gray-50">
                        <td className="p-2 font-mono font-semibold">{group.part}</td>
                        <td className="p-2 text-xs">{first.description}</td>
                        <td className="p-2 text-xs text-gray-400">{group.indices.length} lines</td>
                        <td className="p-2"></td>
                        <td className="p-2 text-xs">
                          {first.iop_next_demand ? (
                            <span className="text-xs px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-800">
                              {new Date(first.iop_next_demand + "T00:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })}
                            </span>
                          ) : (
                            <span className="text-gray-300">—</span>
                          )}
                        </td>
                        <td className="p-2 font-semibold">{totalQty}</td>
                        <td className="p-2"></td>
                        <td className="p-2">
                          {first.revisions.length > 1 ? (
                            <select
                              className="border rounded p-1 text-sm"
                              value={first.revision}
                                onChange={(e) => {
                                  // Update revision for all rows in this group
                                for (const idx of group.indices) updateMatched(sortedMatchedIndexed[idx].idx, { revision: e.target.value });
                                }}
                            >
                              {first.revisions.map((r) => (
                                <option key={r} value={r}>Rev {r}</option>
                              ))}
                            </select>
                          ) : (
                            <span>Rev {first.revision}</span>
                          )}
                        </td>
                        <td className="p-2 text-xs text-gray-500">Total: {totalQty}</td>
                        <td className="p-2">
                          <input
                            className="border rounded p-1 w-20 text-sm"
                            type="number"
                            min={0}
                            step={1}
                            value={first.priority}
                            onChange={(e) => {
                              for (const idx of group.indices) updateMatched(sortedMatchedIndexed[idx].idx, { priority: Number(e.target.value || 1) });
                            }}
                          />
                        </td>
                        <td className="p-2">
                          <button
                            className="rounded border px-2 py-1 text-xs text-red-600"
                            onClick={() => {
                              // Remove all in group (reverse order to keep indices valid)
                              const sorted = [...group.indices].sort((a, b) => b - a);
                              for (const idx of sorted) removeMatched(sortedMatchedIndexed[idx].idx);
                            }}
                          >Remove All</button>
                        </td>
                      </tr>,
                      // Sub-rows for each WO
                      ...group.indices.map((idx) => {
                        const entry = sortedMatchedIndexed[idx];
                        const m = entry.m;
                        return (
                          <tr key={`sub-${m.part}-${m.order}-${idx}`} className="border-b bg-white">
                            <td className="p-2 pl-6 text-xs text-gray-400">└</td>
                            <td className="p-2"></td>
                            <td className="p-2 text-xs font-mono">{m.order || "MRP Plan"}</td>
                            <td className="p-2">
                              {(() => {
                                const badge = getWoPhaseBadge(m.order || "");
                                return <span className={badge.className}>{badge.label}</span>;
                              })()}
                            </td>
                            <td className="p-2 text-xs">
                              {m.iop_next_demand ? (
                                <span className="text-xs px-1.5 py-0.5 rounded bg-indigo-50 text-indigo-800">
                                  {new Date(m.iop_next_demand + "T00:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric" })}
                                </span>
                              ) : (
                                <span className="text-gray-300">—</span>
                              )}
                            </td>
                            <td className="p-2 text-gray-500">{m.qty}</td>
                            <td className="p-2 text-gray-500">{m.short}</td>
                            <td className="p-2"></td>
                            <td className="p-2">
                              <input
                                className="border rounded p-1 w-20 text-sm"
                                type="number"
                                min={0}
                                value={m.qty}
                                onChange={(e) => updateMatched(entry.idx, { qty: Number(e.target.value || 0) })}
                              />
                            </td>
                            <td className="p-2"></td>
                            <td className="p-2">
                              <button className="rounded border px-2 py-1 text-xs" onClick={() => removeMatched(entry.idx)}>Remove</button>
                            </td>
                          </tr>
                        );
                      }),
                    ];
                  });
                })()}
              </tbody>
            </table>
          </div>

          <div className="flex gap-3 flex-wrap">
            <button className="rounded bg-black text-white px-4 py-2" onClick={runAnalysis} disabled={loading}>
              {loading ? "Analyzing..." : "Run Multi-BOM Analysis"}
            </button>
            <button
              className="rounded bg-emerald-700 text-white px-4 py-2"
              onClick={checkBuildReadiness}
              disabled={readinessLoading}
            >
              {readinessLoading ? "Checking..." : "⚡ Check Build Readiness"}
            </button>
          </div>
          {error && <div className="text-sm text-red-600">{error}</div>}
          {readinessError && <div className="text-sm text-red-600">{readinessError}</div>}
        </section>
      )}

      {completedMatched.length > 0 && (
        <section className="border border-gray-200 rounded p-3 space-y-2 opacity-75">
          <button className="w-full flex items-center justify-between" onClick={() => setShowCompletedWos((v) => !v)}>
            <h2 className="font-semibold">Completed Work Orders ({completedMatched.length})</h2>
            <span>{showCompletedWos ? "▾" : "▸"}</span>
          </button>
          {showCompletedWos && (
            <div className="overflow-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left border-b">
                    <th className="p-2">Part Number</th>
                    <th className="p-2">Description</th>
                    <th className="p-2">Work Order</th>
                    <th className="p-2">Ord Qty</th>
                    <th className="p-2">Status</th>
                  </tr>
                </thead>
                <tbody>
                  {completedMatched.map((m, i) => (
                    <tr key={`${m.part}-${m.order}-${i}`} className="border-b">
                      <td className="p-2 font-mono">{m.part}</td>
                      <td className="p-2 text-xs">{m.description}</td>
                      <td className="p-2 text-xs font-mono">{m.order || "—"}</td>
                      <td className="p-2">{m.qty}</td>
                      <td className="p-2">{woSmtStatusMap.get((m.order || "").trim()) || "Complete"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>
      )}

      {/* Build Readiness Results */}
      {readiness && (
        <section className="space-y-4">
          <div className="flex items-center justify-between">
            <h2 className="text-xl font-bold">Build Readiness</h2>
            <button
              className="rounded border px-3 py-1 text-sm text-gray-600 hover:text-red-600 hover:border-red-300"
              onClick={() => { setReadiness(null); setExpandedReadiness(new Set()); }}
            >✕ Clear</button>
          </div>

          {/* Summary cards */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div className="border rounded p-4 bg-green-50">
              <div className="text-3xl font-bold text-green-700">
                {readiness.filter((r) => r.status === "clear-xray").length}
              </div>
              <div className="text-sm text-green-800 mt-1">✅ Clear from XRAY</div>
              <div className="text-xs text-green-600">All parts available in XRAY inventory</div>
            </div>
            <div className="border rounded p-4 bg-blue-50">
              <div className="text-3xl font-bold text-blue-700">
                {readiness.filter((r) => r.status === "clear-with-stockroom").length}
              </div>
              <div className="text-sm text-blue-800 mt-1">📦 Clear with Stockroom</div>
              <div className="text-xs text-blue-600">Needs stockroom parts to complete</div>
            </div>
            <div className="border rounded p-4 bg-red-50">
              <div className="text-3xl font-bold text-red-700">
                {readiness.filter((r) => r.status === "not-clear").length}
              </div>
              <div className="text-sm text-red-800 mt-1">❌ Not Clear</div>
              <div className="text-xs text-red-600">Missing parts even with stockroom</div>
            </div>
          </div>

          {/* Detail rows */}
          <div className="space-y-2">
            {readiness.map((r) => {
              const key = r.run_id || `${r.assembly_pn}__${r.revision}`;
              const isExpanded = expandedReadiness.has(key);
              const statusConfig = {
                "clear-xray": { bg: "bg-green-50", border: "border-green-200", icon: "✅", label: "Clear from XRAY", badge: "bg-green-100 text-green-800" },
                "clear-with-stockroom": { bg: "bg-blue-50", border: "border-blue-200", icon: "📦", label: "Clear with Stockroom", badge: "bg-blue-100 text-blue-800" },
                "not-clear": { bg: "bg-red-50", border: "border-red-200", icon: "❌", label: "Not Clear", badge: "bg-red-100 text-red-800" },
              }[r.status];

              return (
                <div key={key} className={`border rounded ${statusConfig.border}`}>
                  <button
                    className={`w-full flex items-center justify-between p-3 text-left ${statusConfig.bg}`}
                    onClick={() => toggleReadinessExpand(key)}
                  >
                    <div className="flex items-center gap-3 flex-wrap">
                      <span className="text-lg">{statusConfig.icon}</span>
                      <span className="font-mono font-semibold">{r.assembly_pn}</span>
                      <span className="text-sm text-gray-500">Rev {r.revision}</span>
                      {r.run_label && <span className="text-xs text-gray-600">{r.run_label}</span>}
                      {r.description && <span className="text-sm text-gray-500">— {r.description}</span>}
                      <span className={`text-xs px-2 py-0.5 rounded ${statusConfig.badge}`}>{statusConfig.label}</span>
                    </div>
                    <div className="flex items-center gap-4 text-sm text-gray-600">
                      <span>Qty: <b>{r.qty}</b></span>
                      <span>XRAY: <b>{r.max_buildable_xray}</b>/{r.qty}</span>
                      {r.status !== "clear-xray" && (
                        <span>XRAY+SR: <b>{r.max_buildable_combined}</b>/{r.qty}</span>
                      )}
                      {(() => { const sc = r.details.filter(d => d.is_shared_conflict).length; const s = r.details.filter(d => d.is_shared).length; return s > 0 ? <span className={sc > 0 ? "text-red-600" : "text-orange-600"}>{s} shared{sc > 0 ? ` (${sc} short)` : ""}</span> : null; })()}
                      <span>{isExpanded ? "▾" : "▸"}</span>
                    </div>
                  </button>

                  {isExpanded && r.details.length > 0 && (
                    <div className="overflow-auto border-t">
                      <table className="w-full text-sm">
                        <thead>
                          <tr className="text-left border-b bg-gray-50">
                            <th className="p-2">Part</th>
                            <th className="p-2">Required</th>
                            <th className="p-2">XRAY Avail</th>
                            <th className="p-2">Stockroom Avail</th>
                            <th className="p-2">Combined</th>
                            <th className="p-2">Still Short</th>
                            <th className="p-2">On Order</th>
                            <th className="p-2">Type</th>
                          </tr>
                        </thead>
                        <tbody>
                          {r.details.map((d, i) => {
                            const poList = openPOLookup.get(d.part.toUpperCase()) || [];
                            const poTotal = poList.reduce((sum, po) => sum + po.qty_due, 0);
                            const poKey = `readiness||${key}||${d.part}`;
                            const poExpanded = expandedPoRows.has(poKey);
                            return (
                              <>
                                <tr
                                  key={`${d.part}-${i}`}
                                  className={`border-b ${
                                    d.is_floor_stock
                                      ? "bg-purple-50/50"
                                      : d.shortage === 0
                                        ? d.stockroom_available > 0 ? "bg-blue-50/50" : "bg-green-50/50"
                                        : "bg-red-50/50"
                                  }`}
                                >
                                  <td className="p-2 font-mono">{d.part}</td>
                                  <td className="p-2">{d.required}</td>
                                  <td className="p-2">{d.xray_available}</td>
                                  <td className="p-2">{d.stockroom_available > 0 ? d.stockroom_available : "—"}</td>
                                  <td className="p-2">{d.is_floor_stock ? "—" : d.combined_available}</td>
                                  <td className="p-2 font-semibold">{d.is_floor_stock ? "—" : d.shortage > 0 ? d.shortage : "—"}</td>
                                  <td className="p-2">
                                    {d.shortage > 0 && !d.is_floor_stock ? (
                                      poList.length > 0 ? (
                                        <button className="text-xs rounded border px-2 py-0.5 bg-indigo-50 hover:bg-indigo-100" onClick={() => togglePoExpand(poKey)}>
                                          {poTotal} {poExpanded ? "▾" : "▸"}
                                        </button>
                                      ) : (
                                        <button className="text-xs rounded border px-2 py-0.5 bg-red-50 text-red-600 hover:bg-red-100" onClick={() => togglePoExpand(poKey)}>
                                          ⚠ None {poExpanded ? "▾" : "▸"}
                                        </button>
                                      )
                                    ) : "—"}
                                  </td>
                                  <td className="p-2 space-x-1">
                                    <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${d.mount_type === "SMT" ? "bg-blue-100 text-blue-800" : "bg-gray-100 text-gray-600"}`}>{d.mount_type}</span>
                                    {d.is_floor_stock && (
                                      <span className="text-xs px-1.5 py-0.5 rounded bg-purple-100 text-purple-800">🏭 FI</span>
                                    )}
                                    {d.is_shared_conflict && (
                                      <span className="text-xs px-1.5 py-0.5 rounded bg-red-100 text-red-800" title="Shared across assemblies — total demand exceeds total supply">🔴 Shared Short</span>
                                    )}
                                    {d.is_shared && !d.is_shared_conflict && (
                                      <span className="text-xs px-1.5 py-0.5 rounded bg-orange-100 text-orange-800" title="Shared across multiple assemblies — supply currently covers demand">⚠️ Shared</span>
                                    )}
                                  </td>
                                </tr>
                                {poExpanded && (poList.length > 0 ? poList.map((po, poIdx) => (
                                  <tr key={`${poKey}-${po.po}-${po.line}-${poIdx}`} className="border-b bg-indigo-50/70 text-xs">
                                    <td className="p-2" colSpan={10}>
                                      <span className="font-semibold mr-4">Vendor: {po.vendor || "—"}</span>
                                      <span className="mr-4">PO: {po.po}{po.line ? ` / Line ${po.line}` : ""}</span>
                                      <span className="mr-4">Qty Due: {po.qty_due}</span>
                                      <span className="mr-4">Ordered: {po.qty_ordered}</span>
                                      <span className="mr-4">Received: {po.qty_received}</span>
                                      <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                      <span className="mr-4">Orig Promise: {po.original_promise || "—"}</span>
                                      <span>Description: {po.description || "—"}</span>
                                    </td>
                                  </tr>
                                )) : (
                                  <tr key={`${poKey}-none`} className="border-b bg-red-50/70 text-xs">
                                    <td className="p-2 text-red-600 font-medium" colSpan={10}>⚠️ Part currently has not been ordered</td>
                                  </tr>
                                ))}
                              </>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}

                  {isExpanded && r.details.length === 0 && (
                    <div className="p-3 text-sm text-gray-500 border-t">All parts available — no shortages.</div>
                  )}
                </div>
              );
            })}
          </div>
        </section>
      )}

      {activeMatched.length === 0 && mpsRows.length > 0 && (
        <div className="border border-amber-400 bg-amber-50 rounded p-3 text-sm text-amber-800">
          No board assemblies from the MPS matched your BOM database. Make sure your BOMs are imported in the Data Manager.
        </div>
      )}

      {/* Analysis results — reused from multi-bom */}
      {result && (
        <section className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-5 gap-3">
            <div className="border rounded p-3">Requested Runs: <b>{result.summary.requested_runs}</b></div>
            <div className="border rounded p-3">Requested Boards: <b>{result.summary.requested_boards}</b></div>
            <div className="border rounded p-3">Individually Complete: <b>{result.summary.individually_completable_runs}</b></div>
            <div className="border rounded p-3">Recommended Complete Runs: <b>{result.summary.recommended_complete_runs}</b></div>
            <div className="border rounded p-3">Recommended Boards: <b>{result.summary.recommended_boards}</b></div>
          </div>

          <div className="flex gap-2">
            <button className="rounded border px-3 py-1" onClick={exportCsv}>Export CSV</button>
            <div className="text-sm text-gray-600">Source: {result.source}</div>
          </div>

          {result.errors && result.errors.length > 0 && (
            <div className="border border-amber-400 bg-amber-50 rounded p-3 text-sm text-amber-800">
              <strong>Warnings:</strong>
              <ul className="mt-1 list-disc list-inside">
                {result.errors.map((e, i) => <li key={i}>{e}</li>)}
              </ul>
            </div>
          )}

          <div className="border rounded p-3 space-y-2">
            <h2 className="font-semibold">Plan Results</h2>
            <div className="overflow-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left border-b">
                    <th className="p-2">Assembly</th>
                    <th className="p-2">Rev</th>
                    <th className="p-2">Target Qty</th>
                    <th className="p-2">Priority</th>
                    <th className="p-2">Can Complete</th>
                    <th className="p-2">Max Buildable</th>
                    <th className="p-2">Recommended</th>
                    <th className="p-2">Shortages</th>
                  </tr>
                </thead>
                <tbody>
                  {result.plans.map((p, i) => (
                    <tr key={p.run_id || `${p.assembly_pn}-${p.revision}-${i}`} className="border-b">
                      <td className="p-2">
                        <div>{p.assembly_pn}</div>
                        {p.run_label && <div className="text-xs text-gray-500">{p.run_label}</div>}
                      </td>
                      <td className="p-2">{p.revision}</td>
                      <td className="p-2">{p.qty}</td>
                      <td className="p-2">{p.priority}</td>
                      <td className="p-2">{p.can_complete ? "YES" : "NO"}</td>
                      <td className="p-2">{p.max_buildable}</td>
                      <td className="p-2">{p.in_recommended ? "YES" : "NO"}</td>
                      <td className="p-2">{p.shortages.length}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="border rounded p-3 space-y-3">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <h2 className="font-semibold">Top Shortage Parts</h2>
              <div className="flex items-center gap-2">
                <button className="rounded border px-3 py-1 text-sm" onClick={exportShortagesCsv}>Export CSV</button>
                <div className="flex border rounded overflow-hidden text-sm">
                  <button
                    className={`px-3 py-1 ${shortageView === "by-bom" ? "bg-black text-white" : "bg-white text-black"}`}
                    onClick={() => setShortageView("by-bom")}
                  >By BOM Run</button>
                  <button
                    className={`px-3 py-1 ${shortageView === "cross-bom" ? "bg-black text-white" : "bg-white text-black"}`}
                    onClick={() => setShortageView("cross-bom")}
                  >Cross-BOM</button>
                </div>
              </div>
            </div>

            {result.top_shortages.length === 0 ? (
              <p className="text-sm text-gray-600">None</p>
            ) : shortageView === "by-bom" ? (
              <div className="space-y-2">
                {shortagesByBom.map((group) => {
                  const isCollapsed = collapsedBoms.has(group.bom_run);
                  const totalShortage = group.shortages.reduce((s, x) => s + x.shortage, 0);
                  const isHighPriority = group.priority >= 3;
                  return (
                    <div key={group.bom_run} className="border rounded">
                      <button
                        className={`w-full flex items-center justify-between p-3 text-left text-sm font-medium ${isHighPriority ? "bg-red-50" : "bg-yellow-50"}`}
                        onClick={() => toggleBomCollapse(group.bom_run)}
                      >
                        <span className="flex items-center gap-2">
                          <span className={`inline-block w-2 h-2 rounded-full ${isHighPriority ? "bg-red-500" : "bg-yellow-500"}`} />
                          <span>{group.bom_run}</span>
                          <span className="text-gray-500">— Priority: {group.priority}</span>
                          <span className="text-gray-500">— {group.shortages.length} shortage{group.shortages.length !== 1 ? "s" : ""} (total: {totalShortage})</span>
                        </span>
                        <span>{isCollapsed ? "▸" : "▾"}</span>
                      </button>
                      {!isCollapsed && (
                        <div className="overflow-auto">
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="text-left border-b">
                                <th className="p-2">Part</th>
                                <th className="p-2">Required</th>
                                <th className="p-2">Available</th>
                                <th className="p-2">Shortage</th>
                                <th className="p-2">On Order</th>
                              </tr>
                            </thead>
                            <tbody>
                              {group.shortages.map((s, i) => {
                                const poList = openPOLookup.get(s.part.toUpperCase()) || [];
                                const poTotal = poList.reduce((sum, po) => sum + po.qty_due, 0);
                                const poKey = `by-bom||${group.bom_run}||${s.part}`;
                                const poExpanded = expandedPoRows.has(poKey);
                                return (
                                  <>
                                    <tr key={`${s.part}-${i}`} className={`border-b ${isHighPriority ? "bg-red-50/50" : "bg-yellow-50/50"}`}>
                                      <td className="p-2 font-mono">{s.part}</td>
                                      <td className="p-2">{s.required}</td>
                                      <td className="p-2">{s.available}</td>
                                      <td className="p-2">{s.shortage}</td>
                                      <td className="p-2">
                                        {poList.length > 0 ? (
                                          <button className="text-xs rounded border px-2 py-0.5 bg-indigo-50 hover:bg-indigo-100" onClick={() => togglePoExpand(poKey)}>
                                            {poTotal} {poExpanded ? "▾" : "▸"}
                                          </button>
                                        ) : (
                                          <button className="text-xs rounded border px-2 py-0.5 bg-red-50 text-red-600 hover:bg-red-100" onClick={() => togglePoExpand(poKey)}>
                                            ⚠ None {poExpanded ? "▾" : "▸"}
                                          </button>
                                        )}
                                      </td>
                                    </tr>
                                    {poExpanded && (poList.length > 0 ? poList.map((po, poIdx) => (
                                      <tr key={`${poKey}-${po.po}-${po.line}-${poIdx}`} className="border-b bg-indigo-50/70 text-xs">
                                        <td className="p-2" colSpan={5}>
                                          <span className="font-semibold mr-4">Vendor: {po.vendor || "—"}</span>
                                          <span className="mr-4">PO: {po.po}{po.line ? ` / Line ${po.line}` : ""}</span>
                                          <span className="mr-4">Qty Due: {po.qty_due}</span>
                                          <span className="mr-4">Ordered: {po.qty_ordered}</span>
                                          <span className="mr-4">Received: {po.qty_received}</span>
                                          <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                          <span className="mr-4">Orig Promise: {po.original_promise || "—"}</span>
                                          <span>Description: {po.description || "—"}</span>
                                        </td>
                                      </tr>
                                    )) : (
                                      <tr key={`${poKey}-none`} className="border-b bg-red-50/70 text-xs">
                                        <td className="p-2 text-red-600 font-medium" colSpan={5}>⚠️ Part currently has not been ordered</td>
                                      </tr>
                                    ))}
                                  </>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="overflow-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left border-b">
                      <th className="p-2">Part</th>
                      <th className="p-2">Total Shortage</th>
                      <th className="p-2">BOM Runs</th>
                      <th className="p-2">On Order</th>
                    </tr>
                  </thead>
                  <tbody>
                    {crossBomShortages.map((s) => {
                      const isHighPriority = s.maxPriority >= 3;
                      const poList = openPOLookup.get(s.part.toUpperCase()) || [];
                      const poTotal = poList.reduce((sum, po) => sum + po.qty_due, 0);
                      const poKey = `cross-bom||${s.part}`;
                      const poExpanded = expandedPoRows.has(poKey);
                      return (
                        <>
                          <tr key={s.part} className={`border-b ${isHighPriority ? "bg-red-50" : "bg-yellow-50"}`}>
                            <td className="p-2 font-mono">{s.part}</td>
                            <td className="p-2">
                              <span className="flex items-center gap-2">
                                {s.totalShortage}
                                <span className={`text-xs px-1.5 py-0.5 rounded ${isHighPriority ? "bg-red-200 text-red-800" : "bg-yellow-200 text-yellow-800"}`}>
                                  {isHighPriority ? "HIGH" : "LOW"}
                                </span>
                              </span>
                            </td>
                            <td className="p-2">{s.bomRuns.join(", ")}</td>
                            <td className="p-2">
                              {poList.length > 0 ? (
                                <button className="text-xs rounded border px-2 py-0.5 bg-indigo-50 hover:bg-indigo-100" onClick={() => togglePoExpand(poKey)}>
                                  {poTotal} {poExpanded ? "▾" : "▸"}
                                </button>
                              ) : (
                                <button className="text-xs rounded border px-2 py-0.5 bg-red-50 text-red-600 hover:bg-red-100" onClick={() => togglePoExpand(poKey)}>
                                  ⚠ None {poExpanded ? "▾" : "▸"}
                                </button>
                              )}
                            </td>
                          </tr>
                          {poExpanded && (poList.length > 0 ? poList.map((po, poIdx) => (
                            <tr key={`${poKey}-${po.po}-${po.line}-${poIdx}`} className="border-b bg-indigo-50/70 text-xs">
                              <td className="p-2" colSpan={4}>
                                <span className="font-semibold mr-4">Vendor: {po.vendor || "—"}</span>
                                <span className="mr-4">PO: {po.po}{po.line ? ` / Line ${po.line}` : ""}</span>
                                <span className="mr-4">Qty Due: {po.qty_due}</span>
                                <span className="mr-4">Ordered: {po.qty_ordered}</span>
                                <span className="mr-4">Received: {po.qty_received}</span>
                                <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                <span className="mr-4">Orig Promise: {po.original_promise || "—"}</span>
                                <span>Description: {po.description || "—"}</span>
                              </td>
                            </tr>
                          )) : (
                            <tr key={`${poKey}-none`} className="border-b bg-red-50/70 text-xs">
                              <td className="p-2 text-red-600 font-medium" colSpan={4}>⚠️ Part currently has not been ordered</td>
                            </tr>
                          ))}
                        </>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {result.top_shortages.length > 0 && (
            <div className="pt-2">
              <button
                className="rounded bg-black text-white px-4 py-2"
                onClick={checkStockroom}
                disabled={stockroomLoading}
              >
                {stockroomLoading ? "Checking Stockroom..." : "Check Stockroom for Shortages"}
              </button>
              {stockroomError && <div className="text-sm text-red-600 mt-1">{stockroomError}</div>}
            </div>
          )}

          {stockroomResults && stockroomResults.length > 0 && (
            <div className="border rounded p-3 space-y-2">
              <div className="flex items-center justify-between">
                <h2 className="font-semibold">Stockroom Availability</h2>
                <div className="flex gap-2">
                  <button className="rounded border px-3 py-1 text-sm" onClick={exportStockroomCsv}>Export All CSV</button>
                  <button className="rounded border px-3 py-1 text-sm bg-red-50 text-red-700 hover:bg-red-100" onClick={exportStockroomShortagesOnlyCsv}>Export Shortages Only</button>
                </div>
              </div>
              <div className="overflow-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left border-b">
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("status")}>Status{sortIndicator("status")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("mount_type")}>Type{sortIndicator("mount_type")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("part")}>Part{sortIndicator("part")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("bom_run")}>BOM Run{sortIndicator("bom_run")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("shortage_qty")}>Shortage{sortIndicator("shortage_qty")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("confirmed_qty")}>Stockroom Qty{sortIndicator("confirmed_qty")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("inspection_qty")}>In Inspection{sortIndicator("inspection_qty")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("shared")}>Shared{sortIndicator("shared")}</th>
                      <th className="p-2">On Order</th>
                      <th className="p-2">Locations</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sortedStockroomResults.map((r, i) => {
                      const isFI = floorStockSet.has(r.part);
                      const confirmed = r.confirmed_qty ?? r.stockroom_available;
                      const inInspection = r.inspection_qty ?? 0;
                      const stillShort = !isFI && confirmed < r.shortage_qty;
                      const coveredWithInspection = !isFI && confirmed < r.shortage_qty && (confirmed + inInspection) >= r.shortage_qty;
                      const coverage = isFI
                        ? "bg-purple-50"
                        : confirmed >= r.shortage_qty
                          ? "bg-green-50"
                          : coveredWithInspection
                            ? "bg-orange-50"
                            : confirmed > 0 || inInspection > 0
                              ? "bg-amber-50"
                              : "bg-red-50";
                      const locStr = isFI
                        ? "🏭 Floor Stock (FI)"
                        : r.stockroom_locations.length > 0
                          ? r.stockroom_locations.map((l) => `${l.qty} @ ${l.bin_location} (${l.lot_number})`).join(", ")
                          : "Not in stockroom";
                      const poList = openPOLookup.get(r.part.toUpperCase()) || [];
                      const poTotal = poList.reduce((sum, po) => sum + po.qty_due, 0);
                      const poKey = `stockroom||${r.bom_run}||${r.part}`;
                      const poExpanded = expandedPoRows.has(poKey);
                      return (
                        <Fragment key={`${r.part}-${r.bom_run}-${i}`}>
                          <tr className={`border-b ${coverage}`}>
                            <td className="p-2 text-center">
                              {isFI ? "🏭" : confirmed >= r.shortage_qty ? "🟢" : coveredWithInspection ? "🟠" : confirmed > 0 || inInspection > 0 ? "🟡" : "🔴"}
                            </td>
                            <td className="p-2 text-center">
                              <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${(r.mount_type || "TH") === "SMT" ? "bg-blue-100 text-blue-800" : "bg-gray-100 text-gray-700"}`}>
                                {r.mount_type || "TH"}
                              </span>
                            </td>
                            <td className="p-2 font-mono">
                              {r.part}
                              {isFI && <span className="ml-2 text-xs px-1.5 py-0.5 rounded bg-purple-100 text-purple-800">FI</span>}
                            </td>
                            <td className="p-2">{r.bom_run}</td>
                            <td className="p-2">{isFI ? "—" : r.shortage_qty}</td>
                            <td className="p-2">{isFI ? "—" : confirmed}</td>
                            <td className="p-2">{isFI ? "—" : inInspection > 0 ? <span className="text-orange-600 font-medium">{inInspection} ⚠</span> : "—"}</td>
                            <td className="p-2">
                              {(() => {
                                const shared = sharedPartsInfo.get(r.part);
                                if (!shared || isFI) return "—";
                                const contested = shared.totalShortage > shared.stockroomAvail;
                                return (
                                  <span title={`Needed by ${shared.bomRuns.size} BOMs — Total need: ${shared.totalShortage}, Stockroom: ${shared.stockroomAvail}`}>
                                    <span className={`text-xs px-1.5 py-0.5 rounded ${contested ? "bg-red-200 text-red-800" : "bg-blue-100 text-blue-800"}`}>
                                      {shared.bomRuns.size} BOMs{contested ? " ⚠" : ""}
                                    </span>
                                  </span>
                                );
                              })()}
                            </td>
                            <td className="p-2">
                              {stillShort ? (
                                poList.length > 0 ? (
                                  <button className="text-xs rounded border px-2 py-0.5 bg-indigo-50 hover:bg-indigo-100 text-left" onClick={() => togglePoExpand(poKey)}>
                                    <div>{poTotal} {poExpanded ? "▾" : "▸"}</div>
                                    <div className="text-[10px] text-gray-600">{poList[0]?.vendor || "—"} | {poList[0]?.po || "—"} | {poList[0]?.promise_date || "—"}</div>
                                  </button>
                                ) : (
                                  <button className="text-xs rounded border px-2 py-0.5 bg-red-50 text-red-600 hover:bg-red-100" onClick={() => togglePoExpand(poKey)}>
                                    ⚠ None {poExpanded ? "▾" : "▸"}
                                  </button>
                                )
                              ) : "—"}
                            </td>
                            <td className="p-2 text-xs">{locStr}</td>
                          </tr>
                          {poExpanded && stillShort && (poList.length > 0 ? poList.map((po, poIdx) => (
                            <tr key={`${poKey}-${po.po}-${po.line}-${poIdx}`} className="border-b bg-indigo-50/70 text-xs">
                              <td className="p-2" colSpan={10}>
                                <span className="font-semibold mr-4">Vendor: {po.vendor || "—"}</span>
                                <span className="mr-4">PO: {po.po}{po.line ? ` / Line ${po.line}` : ""}</span>
                                <span className="mr-4">Qty Due: {po.qty_due}</span>
                                <span className="mr-4">Ordered: {po.qty_ordered}</span>
                                <span className="mr-4">Received: {po.qty_received}</span>
                                <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                <span className="mr-4">Orig Promise: {po.original_promise || "—"}</span>
                                <span>Description: {po.description || "—"}</span>
                              </td>
                            </tr>
                          )) : (
                            <tr key={`${poKey}-none`} className="border-b bg-red-50/70 text-xs">
                              <td className="p-2 text-red-600 font-medium" colSpan={10}>⚠️ Part currently has not been ordered</td>
                            </tr>
                          ))}
                        </Fragment>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {stockroomResults && stockroomResults.length === 0 && (
            <div className="border rounded p-3">
              <p className="text-sm text-gray-600">No shortage parts to check against stockroom.</p>
            </div>
          )}
        </section>
      )}
    </main>
  );
}
