"use client";

import { Fragment, useEffect, useMemo, useState } from "react";

type BomItem = { assembly_pn: string; revision: string; line_count: number };
type PlanRow = { assembly_pn: string; revision: string; qty: number; priority: number };

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
  assembly_pn: string;
  revision: string;
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
    assembly_pn: string;
    revision: string;
    qty: number;
    priority: number;
    can_complete: boolean;
    max_buildable: number;
    in_recommended: boolean;
    shortages: Array<{ part: string; required: number; available: number; shortage: number }>;
  }>;
  recommended: Array<{ assembly_pn: string; revision: string; qty: number; priority: number }>;
  top_shortages: Array<{ part: string; shortage: number; required: number; available: number; bom_run: string; priority: number }>;
  errors?: string[];
};

export default function MultiBomPage() {
  const [boms, setBoms] = useState<BomItem[]>([]);
  const [rows, setRows] = useState<PlanRow[]>([{ assembly_pn: "", revision: "", qty: 1, priority: 1 }]);
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
  const [openPOs, setOpenPOs] = useState<Array<{ vendor: string; po: string; line: string; item: string; description: string; qty_ordered: number; qty_received: number; qty_due: number; promise_date: string; original_promise: string }>>([]);
  const [expandedPoRows, setExpandedPoRows] = useState<Set<string>>(new Set());

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
      } catch {
        /* silent */
      }
    })();
  }, []);

  const assemblyOptions = useMemo(() => {
    return boms.map((b) => ({ label: `${b.assembly_pn} rev ${b.revision}`, value: `${b.assembly_pn}__${b.revision}` }));
  }, [boms]);

  function toggleBomCollapse(bomKey: string) {
    setCollapsedBoms((prev) => {
      const next = new Set(prev);
      if (next.has(bomKey)) next.delete(bomKey);
      else next.add(bomKey);
      return next;
    });
  }

  // Group shortages by BOM run, sorted by priority weight desc
  const shortagesByBom = useMemo(() => {
    if (!result) return [];
    const groups = new Map<string, { bom_run: string; priority: number; shortages: typeof result.top_shortages }>();
    for (const s of result.top_shortages) {
      if (!groups.has(s.bom_run)) {
        groups.set(s.bom_run, { bom_run: s.bom_run, priority: s.priority, shortages: [] });
      }
      groups.get(s.bom_run)!.shortages.push(s);
    }
    // Sort shortages within each group by shortage desc
    for (const g of groups.values()) {
      g.shortages.sort((a, b) => b.shortage - a.shortage);
    }
    // Sort groups by priority desc
    return Array.from(groups.values()).sort((a, b) => b.priority - a.priority);
  }, [result]);

  // Cross-BOM view: aggregate same part across BOMs
  const crossBomShortages = useMemo(() => {
    if (!result) return [];
    const partMap = new Map<string, { part: string; totalShortage: number; bomRuns: string[]; maxPriority: number }>();
    for (const s of result.top_shortages) {
      if (!partMap.has(s.part)) {
        partMap.set(s.part, { part: s.part, totalShortage: 0, bomRuns: [], maxPriority: 0 });
      }
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
      // Collect ALL shortage parts from individual plans (not top_shortages which is capped at 100)
      const seen = new Set<string>();
      const shortages: Array<{ part: string; shortage: number; bom_run: string }> = [];
      for (const plan of result.plans) {
        const bomRun = `${plan.assembly_pn} rev ${plan.revision}`;
        for (const s of plan.shortages) {
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
      setStockroomResults((data.results || []).map((r: StockroomResult) => ({
        ...r,
        mount_type: smtPartsSet.has(r.part) ? "SMT" as const : "TH" as const,
      })));
    } catch (e: unknown) {
      const message = e instanceof Error ? e.message : "Unknown error";
      setStockroomError(message);
    } finally {
      setStockroomLoading(false);
    }
  }

  function addRow() {
    setRows((prev) => [...prev, { assembly_pn: "", revision: "", qty: 1, priority: 1 }]);
  }

  function removeRow(idx: number) {
    setRows((prev) => prev.filter((_, i) => i !== idx));
  }

  function updateRow(idx: number, next: Partial<PlanRow>) {
    setRows((prev) => prev.map((r, i) => (i === idx ? { ...r, ...next } : r)));
  }

  async function runAnalysis() {
    setLoading(true);
    setError(null);
    setResult(null);
    setStockroomResults(null);
    setStockroomError(null);
    try {
      const plans = rows
        .map((r) => ({ ...r, qty: Math.max(1, Math.floor(Number(r.qty || 0))), priority: Number(r.priority || 1) }))
        .filter((r) => r.assembly_pn && r.revision && r.qty > 0);

      if (plans.length === 0) throw new Error("Add at least one valid assembly row");

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
      const plans = rows
        .map((r) => ({ ...r, qty: Math.max(1, Math.floor(Number(r.qty || 0))), priority: Number(r.priority || 1) }))
        .filter((r) => r.assembly_pn && r.revision && r.qty > 0);

      if (plans.length === 0) throw new Error("Add at least one valid assembly row");

      // Step 1: Run multi-BOM analysis
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
        for (const s of plan.shortages) {
          const key = `${s.part}||${plan.assembly_pn} rev ${plan.revision}`;
          if (!allShortageParts.has(key)) {
            allShortageParts.set(key, { part: s.part, shortage: s.shortage, bom_run: `${plan.assembly_pn} rev ${plan.revision}` });
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
            // Also set part-only key (max across all bom_runs)
            const existing = stockroomMap.get(r.part) || 0;
            if (r.stockroom_available > existing) stockroomMap.set(r.part, r.stockroom_available);
          }
        }
      }

      // Step 3: Build readiness per assembly
      const readinessRows: ReadinessRow[] = [];
      for (const plan of analysisData.plans) {
        const details: ReadinessRow["details"] = [];
        let hasShortageAfterStockroom = false;

        for (const s of plan.shortages) {
          const isFI = floorStockSet.has(s.part);
          const srKey = `${s.part}||${plan.assembly_pn} rev ${plan.revision}`;
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
            is_shared_conflict: false,
            is_shared: false,
          });
        }

        const nonFiDetails = details.filter((d) => !d.is_floor_stock);
        let maxBuildableCombined = plan.max_buildable;
        if (nonFiDetails.length > 0 && plan.qty > 0) {
          let minBoardsCombined = Number.MAX_SAFE_INTEGER;
          for (const d of nonFiDetails) {
            const qtyPerBoard = d.required / plan.qty;
            if (qtyPerBoard > 0) {
              minBoardsCombined = Math.min(minBoardsCombined, Math.floor(d.combined_available / qtyPerBoard));
            }
          }
          if (minBoardsCombined !== Number.MAX_SAFE_INTEGER) {
            maxBuildableCombined = Math.min(plan.qty, Math.max(plan.max_buildable, minBoardsCombined));
          }
        } else if (nonFiDetails.length === 0 && details.length > 0) {
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

        readinessRows.push({
          assembly_pn: plan.assembly_pn,
          revision: plan.revision,
          qty: plan.qty,
          status,
          max_buildable_xray: plan.max_buildable,
          max_buildable_combined: maxBuildableCombined,
          shortage_parts_xray: plan.shortages.length,
          shortage_parts_after_stockroom: nonFiShortageCount,
          details: details.sort((a, b) => b.shortage - a.shortage),
        });
      }

      // Detect shared part conflicts across assemblies
      if (readinessRows.length > 1) {
        // Sum total demand per part across all assemblies
        const totalDemand = new Map<string, number>();
        const totalSupply = new Map<string, number>(); // xray + stockroom (first seen)
        for (const row of readinessRows) {
          for (const d of row.details) {
            if (d.is_floor_stock) continue;
            totalDemand.set(d.part, (totalDemand.get(d.part) || 0) + d.required);
            if (!totalSupply.has(d.part)) {
              totalSupply.set(d.part, d.combined_available);
            }
          }
        }
        // Flag shared parts and shared-short parts
        const sharedParts = new Set<string>();
        const conflictParts = new Set<string>();
        for (const [part, demand] of totalDemand) {
          // Check if part appears in multiple assemblies
          let count = 0;
          for (const row of readinessRows) {
            if (row.details.some((d) => d.part === part && !d.is_floor_stock)) count++;
          }
          if (count > 1) {
            sharedParts.add(part);
            const supply = totalSupply.get(part) || 0;
            if (demand > supply) conflictParts.add(part);
          }
        }
        // Mark affected details
        for (const row of readinessRows) {
          for (const d of row.details) {
            if (conflictParts.has(d.part)) { d.is_shared_conflict = true; d.is_shared = true; }
            else if (sharedParts.has(d.part)) { d.is_shared = true; }
          }
        }
      }

      const statusOrder: Record<ReadinessStatus, number> = { "clear-xray": 0, "clear-with-stockroom": 1, "not-clear": 2 };
      readinessRows.sort((a, b) => statusOrder[a.status] - statusOrder[b.status]);
      setReadiness(readinessRows);
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
  }, [stockroomResults, stockroomSort, sharedPartsInfo]);

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
    const header = ["assembly_pn", "revision", "qty", "priority", "can_complete", "max_buildable", "recommended"];
    const lines = result.plans.map((p) => [p.assembly_pn, p.revision, p.qty, p.priority, p.can_complete ? "YES" : "NO", p.max_buildable, p.in_recommended ? "YES" : "NO"]);
    downloadCsv("multi-bom-analysis.csv", [header, ...lines].map((r) => r.join(",")).join("\n"));
  }

  function exportShortagesCsv() {
    if (!result) return;
    const header = ["part", "shortage", "required", "available", "bom_run", "priority"];
    const lines = result.top_shortages.map((s) => [s.part, s.shortage, s.required, s.available, s.bom_run, s.priority]);
    downloadCsv("top-shortage-parts.csv", [header, ...lines].map((r) => r.join(",")).join("\n"));
  }

  function buildSourcingHref(): string {
    if (!result || result.top_shortages.length === 0) return "/sourcing";
    // Dedupe by part: sum shortages across BOM runs so Sourcing sees one line per PN.
    const totals = new Map<string, number>();
    for (const s of result.top_shortages) {
      if (!s.part || s.shortage <= 0) continue;
      totals.set(s.part, (totals.get(s.part) || 0) + s.shortage);
    }
    const parts: string[] = [];
    // Cap at 200 parts in URL to keep it reasonable.
    let count = 0;
    for (const [pn, qty] of totals) {
      if (count >= 200) break;
      parts.push(`${encodeURIComponent(pn)}:${qty}`);
      count++;
    }
    const prefill = parts.join(",");
    const source = encodeURIComponent("Multi-BOM shortages");
    return `/sourcing?prefill=${prefill}&source=${source}`;
  }

  function exportStockroomCsv() {
    if (!stockroomResults) return;
    const header = ["part", "mount_type", "bom_run", "shortage", "confirmed_stockroom_qty", "in_inspection", "shared_boms", "contested", "locations"];
    const lines = stockroomResults.map((s) => {
      const shared = sharedPartsInfo.get(s.part);
      const conf = s.confirmed_qty ?? s.stockroom_available;
      const insp = s.inspection_qty ?? 0;
      return [
        s.part,
        s.mount_type || "TH",
        s.bom_run,
        s.shortage_qty,
        conf,
        insp > 0 ? insp : "",
        shared ? shared.bomRuns.size : "",
        shared && shared.totalShortage > shared.stockroomAvail ? "YES" : "",
        `"${s.stockroom_locations.map((l) => `${l.qty} @ ${l.bin_location} (${l.lot_number})`).join("; ")}"`,
      ];
    });
    downloadCsv("stockroom-availability.csv", [header, ...lines].map((r) => r.join(",")).join("\n"));
  }

  return (
    <main className="max-w-5xl mx-auto p-6 space-y-6 w-full">
      <h1 className="text-2xl font-bold">Multi-BOM Analysis</h1>
      <p className="text-sm text-gray-600">Complete-run-first optimizer: prioritize full uninterrupted runs, then boards, then priority. Priority Weight is only a tie-breaker when material is limited.</p>

      <section className="border rounded p-3 space-y-3">
        <div className="flex items-center justify-between">
          <h2 className="font-semibold">Plan Inputs</h2>
          <button className="rounded border px-3 py-1" onClick={addRow}>+ Add Assembly</button>
        </div>

        <div className="space-y-2">
          <div className="hidden md:grid md:grid-cols-12 gap-2 text-xs text-gray-600 px-1">
            <div className="md:col-span-6">Assembly (PN + Rev)</div>
            <div className="md:col-span-2">Target Qty</div>
            <div className="md:col-span-2" title="Higher number = higher priority. Used as tie-breaker when material is limited.">Priority Weight ℹ️</div>
            <div className="md:col-span-2">Action</div>
          </div>
          {rows.map((r, idx) => (
            <div key={idx} className="grid grid-cols-1 md:grid-cols-12 gap-2 items-center border rounded p-2">
              <select
                className="border rounded p-2 md:col-span-6"
                value={r.assembly_pn && r.revision ? `${r.assembly_pn}__${r.revision}` : ""}
                onChange={(e) => {
                  const value = e.target.value;
                  if (!value) return updateRow(idx, { assembly_pn: "", revision: "" });
                  const [assembly_pn, revision] = value.split("__");
                  updateRow(idx, { assembly_pn, revision });
                }}
              >
                <option value="">Select Assembly</option>
                {assemblyOptions.map((o) => (
                  <option key={o.value} value={o.value}>{o.label}</option>
                ))}
              </select>

              <input
                className="border rounded p-2 md:col-span-2"
                type="number"
                min={1}
                value={r.qty}
                onChange={(e) => updateRow(idx, { qty: Number(e.target.value || 1) })}
                placeholder="Target Qty"
              />

              <input
                className="border rounded p-2 md:col-span-2"
                type="number"
                min={0}
                step={1}
                value={r.priority}
                onChange={(e) => updateRow(idx, { priority: Number(e.target.value || 1) })}
                placeholder="Priority (higher = first)"
              />

              <button className="rounded border px-3 py-2 md:col-span-2" onClick={() => removeRow(idx)} disabled={rows.length === 1}>Remove</button>
            </div>
          ))}
        </div>

        <div className="flex gap-3 flex-wrap">
          <button className="rounded bg-black text-white px-4 py-2" onClick={runAnalysis}>
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

      {openPOs.length > 0 && (
        <div className="text-sm text-gray-500 border rounded p-2">
          📋 {openPOs.length} Open PO lines loaded. <span className="text-xs">(Manage uploads in <a href="/data" className="underline">Data Manager</a>)</span>
        </div>
      )}

      {/* Build Readiness Results */}
      {readiness && (
        <section className="space-y-4">
          <h2 className="text-xl font-bold">Build Readiness</h2>

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

          <div className="space-y-2">
            {readiness.map((r) => {
              const key = `${r.assembly_pn}__${r.revision}`;
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
                              <Fragment key={`${poKey}-${i}`}>
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
                                  <td className="p-2">
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
                                    <td className="p-2" colSpan={8}>
                                      <span className="font-semibold mr-4">Vendor: {po.vendor || "—"}</span>
                                      <span className="mr-4">PO: {po.po}{po.line ? ` / Line ${po.line}` : ""}</span>
                                      <span className="mr-4">Qty Due: {po.qty_due}</span>
                                      <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                      <span>Orig Promise: {po.original_promise || "—"}</span>
                                    </td>
                                  </tr>
                                )) : (
                                  <tr key={`${poKey}-none`} className="border-b bg-red-50/70 text-xs">
                                    <td className="p-2 text-red-600 font-medium" colSpan={8}>⚠️ Part currently has not been ordered</td>
                                  </tr>
                                ))}
                              </Fragment>
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
                    <th className="p-2">Priority Weight</th>
                    <th className="p-2">Can Complete</th>
                    <th className="p-2">Max Buildable Qty</th>
                    <th className="p-2">Recommended</th>
                    <th className="p-2">Shortages</th>
                  </tr>
                </thead>
                <tbody>
                  {result.plans.map((p) => (
                    <tr key={`${p.assembly_pn}-${p.revision}`} className="border-b">
                      <td className="p-2">{p.assembly_pn}</td>
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
                <a
                  href={buildSourcingHref()}
                  className="rounded bg-black text-white px-3 py-1 text-sm hover:bg-gray-800"
                  title="Open these shortage parts in the Sourcing Optimizer"
                >
                  📦 Check sourcing
                </a>
                <button className="rounded border px-3 py-1 text-sm" onClick={exportShortagesCsv}>Export CSV</button>
              <div className="flex border rounded overflow-hidden text-sm">
                <button
                  className={`px-3 py-1 ${shortageView === "by-bom" ? "bg-black text-white" : "bg-white text-black"}`}
                  onClick={() => setShortageView("by-bom")}
                >
                  By BOM Run
                </button>
                <button
                  className={`px-3 py-1 ${shortageView === "cross-bom" ? "bg-black text-white" : "bg-white text-black"}`}
                  onClick={() => setShortageView("cross-bom")}
                >
                  Cross-BOM
                </button>
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
                                  <Fragment key={`${poKey}-${i}`}>
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
                                          <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                          <span>Orig Promise: {po.original_promise || "—"}</span>
                                        </td>
                                      </tr>
                                    )) : (
                                      <tr key={`${poKey}-none`} className="border-b bg-red-50/70 text-xs">
                                        <td className="p-2 text-red-600 font-medium" colSpan={5}>⚠️ Part currently has not been ordered</td>
                                      </tr>
                                    ))}
                                  </Fragment>
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
                        <Fragment key={s.part}>
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
                                <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                <span>Orig Promise: {po.original_promise || "—"}</span>
                              </td>
                            </tr>
                          )) : (
                            <tr key={`${poKey}-none`} className="border-b bg-red-50/70 text-xs">
                              <td className="p-2 text-red-600 font-medium" colSpan={4}>⚠️ Part currently has not been ordered</td>
                            </tr>
                          ))}
                        </Fragment>
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
                <button className="rounded border px-3 py-1 text-sm" onClick={exportStockroomCsv}>Export CSV</button>
              </div>
              <div className="overflow-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left border-b">
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("status")}>Status{sortIndicator("status")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("part")}>Part{sortIndicator("part")}</th>
                      <th className="p-2 cursor-pointer select-none hover:bg-gray-100" onClick={() => toggleStockroomSort("mount_type")}>Type{sortIndicator("mount_type")}</th>
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
                    {sortedStockroomResults
                      .map((r, i) => {
                        const isFI = floorStockSet.has(r.part);
                        const confirmed = r.confirmed_qty ?? r.stockroom_available;
                        const inInspection = r.inspection_qty ?? 0;
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
                            ? r.stockroom_locations
                                .map((l) => `${l.qty} @ ${l.bin_location} (${l.lot_number})`)
                                .join(", ")
                            : "Not in stockroom";
                        const poList = openPOLookup.get(r.part.toUpperCase()) || [];
                        const poTotal = poList.reduce((sum, po) => sum + po.qty_due, 0);
                        const poKey = `stockroom||${r.bom_run}||${r.part}`;
                        const poExpanded = expandedPoRows.has(poKey);
                        const shared = sharedPartsInfo.get(r.part);
                        return (
                          <Fragment key={`${r.part}-${r.bom_run}-${i}`}>
                            <tr key={`${r.part}-${r.bom_run}-${i}`} className={`border-b ${coverage}`}>
                              <td className="p-2 text-center">
                                {isFI ? "🏭" : confirmed >= r.shortage_qty ? "🟢" : coveredWithInspection ? "🟠" : confirmed > 0 || inInspection > 0 ? "🟡" : "🔴"}
                              </td>
                              <td className="p-2 font-mono">
                                {r.part}
                                {isFI && <span className="ml-2 text-xs px-1.5 py-0.5 rounded bg-purple-100 text-purple-800">FI</span>}
                              </td>
                              <td className="p-2">
                                <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${(r.mount_type || "TH") === "SMT" ? "bg-blue-100 text-blue-800" : "bg-gray-100 text-gray-700"}`}>
                                  {r.mount_type || "TH"}
                                </span>
                              </td>
                              <td className="p-2">{r.bom_run}</td>
                              <td className="p-2">{isFI ? "—" : r.shortage_qty}</td>
                              <td className="p-2">{isFI ? "—" : confirmed}</td>
                              <td className="p-2">{isFI ? "—" : inInspection > 0 ? <span className="text-orange-600 font-medium">{inInspection} ⚠</span> : "—"}</td>
                              <td className="p-2">
                                {(() => {
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
                                {isFI ? "—" : poList.length > 0 ? (
                                  <button className="text-xs rounded border px-2 py-0.5 bg-indigo-50 hover:bg-indigo-100" onClick={() => togglePoExpand(poKey)}>
                                    {poTotal} {poExpanded ? "▾" : "▸"}
                                  </button>
                                ) : (
                                  <button className="text-xs rounded border px-2 py-0.5 bg-red-50 text-red-600 hover:bg-red-100" onClick={() => togglePoExpand(poKey)}>
                                    ⚠ None {poExpanded ? "▾" : "▸"}
                                  </button>
                                )}
                              </td>
                              <td className="p-2 text-xs">{locStr}</td>
                            </tr>
                            {poExpanded && !isFI && (poList.length > 0 ? poList.map((po, poIdx) => (
                              <tr key={`${poKey}-${po.po}-${po.line}-${poIdx}`} className="border-b bg-indigo-50/70 text-xs">
                                <td className="p-2" colSpan={10}>
                                  <span className="font-semibold mr-4">Vendor: {po.vendor || "—"}</span>
                                  <span className="mr-4">PO: {po.po}{po.line ? ` / Line ${po.line}` : ""}</span>
                                  <span className="mr-4">Qty Due: {po.qty_due}</span>
                                  <span className="mr-4">Promise: {po.promise_date || "—"}</span>
                                  <span>Orig Promise: {po.original_promise || "—"}</span>
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
