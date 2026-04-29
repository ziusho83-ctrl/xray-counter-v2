"use client";

import { type DragEvent, Suspense, useEffect, useMemo, useState } from "react";
import { useSearchParams } from "next/navigation";
import * as XLSX from "xlsx";

type Offer = {
  source: string;
  unit_price: number;
  stock: number;
  lead_time_days: number;
  moq: number;
};

type PlanRow = {
  requested_pn: string;
  picked_pn: string;
  qty: number;
  status: "ok" | "partial_stock" | "no_source";
  substituted_from: string;
  sub_reason: string;
  best: Offer | null;
  alternates: Offer[];
  note: string;
  tried_parts: string[];
  ext_price: number;
};

type OptimizeResult = {
  rows: PlanRow[];
  summary: {
    lines_total: number;
    fully_sourced: number;
    partial_stock: number;
    no_source: number;
    substituted: number;
    total_cost: number;
    critical_lead_days: number;
  };
};

type ParsedBom = { part_number: string; qty: number }[];

const SAMPLE_BOM: ParsedBom = [
  { part_number: "RC0603FR-0710KL", qty: 2500 },
  { part_number: "CL10A475KP8NNNC", qty: 1200 },
  { part_number: "STM32F103C8T6", qty: 50 },
  { part_number: "LM358DR", qty: 200 },
  { part_number: "CRCW06031K00FKEA", qty: 800 },
  { part_number: "ATMEGA328P-PU", qty: 75 },
  { part_number: "GRM188R61A475KE15", qty: 500 },
  { part_number: "MCP6002-I/SN", qty: 100 },
  { part_number: "STM32F103CBT6", qty: 20 },
  { part_number: "TEST-EOL-CAP-001", qty: 10 }, // intentional EOL example to exercise no-source path
];

function parseBomFromWorkbook(data: ArrayBuffer): ParsedBom {
  const wb = XLSX.read(data, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
  const out: ParsedBom = [];
  for (const raw of rows) {
    // Normalize keys (case-insensitive)
    const map: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(raw)) map[k.toLowerCase().trim()] = v;
    const pn = String(
      map["part_number"] ?? map["part number"] ?? map["pn"] ?? map["mpn"] ?? "",
    ).trim();
    const qtyRaw = map["qty"] ?? map["quantity"] ?? map["q"] ?? "";
    const qty = Number(qtyRaw);
    if (!pn || !Number.isFinite(qty) || qty <= 0) continue;
    out.push({ part_number: pn, qty: Math.floor(qty) });
  }
  return out;
}

export default function SourcingPage() {
  return (
    <Suspense fallback={<main className="max-w-6xl mx-auto p-6 text-sm text-gray-600">Loading…</main>}>
      <SourcingPageInner />
    </Suspense>
  );
}

function SourcingPageInner() {
  const searchParams = useSearchParams();
  const [bom, setBom] = useState<ParsedBom>([]);
  const [fileName, setFileName] = useState<string>("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<OptimizeResult | null>(null);
  const [expanded, setExpanded] = useState<Set<number>>(new Set());
  const [dragOver, setDragOver] = useState(false);
  const [prefillBanner, setPrefillBanner] = useState<string>("");

  // Accept ?prefill=PN1:qty1,PN2:qty2 from URL (e.g. linked from Multi-BOM shortages)
  useEffect(() => {
    const prefill = searchParams.get("prefill");
    const source = searchParams.get("source") || "";
    if (!prefill) return;
    const parsed: ParsedBom = [];
    for (const entry of prefill.split(",")) {
      const [pn, qtyRaw] = entry.split(":");
      if (!pn) continue;
      const qty = Number(qtyRaw);
      if (!Number.isFinite(qty) || qty <= 0) continue;
      parsed.push({ part_number: pn.trim(), qty: Math.floor(qty) });
    }
    if (parsed.length > 0) {
      setBom(parsed);
      setFileName(source ? `prefill from ${source}` : "prefill from URL");
      setPrefillBanner(
        `📦 Pre-loaded ${parsed.length} part${parsed.length === 1 ? "" : "s"}${
          source ? ` from ${source}` : ""
        }. Click “Run sourcing optimizer” below.`,
      );
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const hasBom = bom.length > 0;

  function toggleExpand(i: number) {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(i)) next.delete(i);
      else next.add(i);
      return next;
    });
  }

  async function handleFile(file: File) {
    setError(null);
    setResult(null);
    setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseBomFromWorkbook(buf);
      if (parsed.length === 0) {
        setError("No valid rows found. Need columns: part_number, qty.");
        setBom([]);
        return;
      }
      setBom(parsed);
    } catch (e: any) {
      setError(e?.message || "Failed to read file.");
      setBom([]);
    }
  }

  function onDrop(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) void handleFile(file);
  }

  function loadSample() {
    setError(null);
    setResult(null);
    setFileName("sample_bom.csv");
    setBom(SAMPLE_BOM);
  }

  async function runOptimize() {
    if (bom.length === 0) return;
    setLoading(true);
    setError(null);
    setResult(null);
    setExpanded(new Set());
    try {
      const res = await fetch("/api/sourcing/optimize", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ bom }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data?.error || "Optimize failed");
      setResult(data as OptimizeResult);
    } catch (e: any) {
      setError(e?.message || "Unknown error");
    } finally {
      setLoading(false);
    }
  }

  function downloadPlanCsv() {
    if (!result) return;
    const headers = [
      "requested_pn",
      "picked_pn",
      "qty",
      "status",
      "substituted_from",
      "sub_reason",
      "source",
      "unit_price",
      "ext_price",
      "lead_time_days",
      "moq",
      "stock",
      "note",
    ];
    const lines = [headers.join(",")];
    for (const r of result.rows) {
      const cells = [
        r.requested_pn,
        r.picked_pn,
        String(r.qty),
        r.status,
        r.substituted_from,
        r.sub_reason,
        r.best?.source ?? "",
        r.best ? r.best.unit_price.toFixed(4) : "",
        r.best ? r.ext_price.toFixed(2) : "",
        r.best ? String(r.best.lead_time_days) : "",
        r.best ? String(r.best.moq) : "",
        r.best ? String(r.best.stock) : "",
        r.note,
      ].map((c) => {
        const s = String(c ?? "");
        return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
      });
      lines.push(cells.join(","));
    }
    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `sourcing-plan-${Date.now()}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  const substitutions = useMemo(
    () => (result ? result.rows.filter((r) => r.substituted_from) : []),
    [result],
  );
  const unsourced = useMemo(
    () => (result ? result.rows.filter((r) => r.status === "no_source") : []),
    [result],
  );

  // Severity: full red only if >25% of lines are unsourced, otherwise amber.
  const unsourcedSeverity: "none" | "warn" | "alert" = useMemo(() => {
    if (!result || unsourced.length === 0) return "none";
    const ratio = unsourced.length / Math.max(result.summary.lines_total, 1);
    return ratio > 0.25 ? "alert" : "warn";
  }, [result, unsourced.length]);

  // Buy List: group sourced rows by vendor for per-vendor cart CSV generation.
  type VendorBucket = {
    source: string;
    lines: Array<{
      picked_pn: string;
      requested_pn: string;
      qty: number;
      unit_price: number;
      ext_price: number;
      lead_time_days: number;
      moq: number;
      substituted_from: string;
    }>;
    line_count: number;
    total: number;
    longest_lead: number;
  };
  const buyList = useMemo<VendorBucket[]>(() => {
    if (!result) return [];
    const map = new Map<string, VendorBucket>();
    for (const r of result.rows) {
      if (!r.best) continue;
      const source = r.best.source;
      let bucket = map.get(source);
      if (!bucket) {
        bucket = {
          source,
          lines: [],
          line_count: 0,
          total: 0,
          longest_lead: 0,
        };
        map.set(source, bucket);
      }
      const effQty = Math.max(r.qty, r.best.moq || 1);
      bucket.lines.push({
        picked_pn: r.picked_pn,
        requested_pn: r.requested_pn,
        qty: effQty,
        unit_price: r.best.unit_price,
        ext_price: r.ext_price,
        lead_time_days: r.best.lead_time_days,
        moq: r.best.moq,
        substituted_from: r.substituted_from,
      });
      bucket.line_count += 1;
      bucket.total += r.ext_price;
      if (r.best.lead_time_days > bucket.longest_lead) {
        bucket.longest_lead = r.best.lead_time_days;
      }
    }
    return Array.from(map.values()).sort((a, b) => b.total - a.total);
  }, [result]);

  function csvEscape(s: string): string {
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }

  function buildVendorCsv(bucket: VendorBucket): { filename: string; content: string } {
    const src = bucket.source.toLowerCase();
    const lines: string[] = [];
    let filename = `buy-list-${bucket.source.toLowerCase().replace(/[^a-z0-9]+/g, "-")}.csv`;

    if (src.includes("digikey") || src === "digikey") {
      // DigiKey bulk add format: Digi-Key Part Number, Manufacturer Part Number, Quantity, Customer Reference
      lines.push(["Digi-Key Part Number", "Manufacturer Part Number", "Quantity", "Customer Reference"].join(","));
      for (const l of bucket.lines) {
        lines.push(
          [
            "", // Digi-Key PN unknown in mock mode; DigiKey cart resolves by MPN
            csvEscape(l.picked_pn),
            String(l.qty),
            csvEscape(l.substituted_from || l.picked_pn),
          ].join(","),
        );
      }
      filename = "digikey-cart.csv";
    } else if (src.includes("mouser")) {
      // Mouser bulk upload: Mouser #, MfrPart #, Qty, CustomerPart #
      lines.push(["Mouser #", "Mfr Part #", "Qty", "Customer Part #"].join(","));
      for (const l of bucket.lines) {
        lines.push(
          [
            "",
            csvEscape(l.picked_pn),
            String(l.qty),
            csvEscape(l.substituted_from || l.picked_pn),
          ].join(","),
        );
      }
      filename = "mouser-cart.csv";
    } else if (src.includes("arrow")) {
      // Arrow basic format
      lines.push(["Mfr Part Number", "Quantity", "Customer Reference"].join(","));
      for (const l of bucket.lines) {
        lines.push(
          [csvEscape(l.picked_pn), String(l.qty), csvEscape(l.substituted_from || l.picked_pn)].join(","),
        );
      }
      filename = "arrow-cart.csv";
    } else if (src.includes("lcsc")) {
      lines.push(["Manufacture Part Number", "Order Qty.", "Customer NO."].join(","));
      for (const l of bucket.lines) {
        lines.push(
          [csvEscape(l.picked_pn), String(l.qty), csvEscape(l.substituted_from || l.picked_pn)].join(","),
        );
      }
      filename = "lcsc-cart.csv";
    } else {
      // Generic PO CSV with full pricing for any other vendor (TI_Direct, etc.)
      lines.push(
        [
          "vendor",
          "mfr_part_number",
          "quantity",
          "unit_price",
          "ext_price",
          "lead_time_days",
          "customer_reference",
        ].join(","),
      );
      for (const l of bucket.lines) {
        lines.push(
          [
            csvEscape(bucket.source),
            csvEscape(l.picked_pn),
            String(l.qty),
            l.unit_price.toFixed(4),
            l.ext_price.toFixed(2),
            String(l.lead_time_days),
            csvEscape(l.substituted_from || l.picked_pn),
          ].join(","),
        );
      }
    }
    return { filename, content: lines.join("\n") };
  }

  function downloadVendorCsv(bucket: VendorBucket) {
    const { filename, content } = buildVendorCsv(bucket);
    const blob = new Blob([content], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  function downloadAllVendorCsvs() {
    // Sequentially trigger downloads (one per vendor). Browser allows multiple
    // auto-downloads on same user gesture.
    for (const bucket of buyList) {
      downloadVendorCsv(bucket);
    }
  }

  return (
    <main className="max-w-6xl mx-auto p-6 space-y-6 w-full">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <h1 className="text-2xl font-bold">Sourcing Optimizer</h1>
          <p className="text-sm text-gray-600">
            Upload a BOM → ranked sourcing plan with automatic manufacturer
            cross-reference substitutions. Mock data mode (v0.2).
          </p>
        </div>
        <span className="text-xs px-2 py-1 rounded bg-yellow-100 text-yellow-800 border border-yellow-300">
          MOCK DATA — no live pricing
        </span>
      </div>

      {prefillBanner && (
        <div className="border rounded-lg p-3 bg-indigo-50 border-indigo-200 text-sm text-indigo-900">
          {prefillBanner}
        </div>
      )}

      {/* Upload card */}
      <section className="border rounded-lg p-4 space-y-3">
        <h2 className="font-semibold">1. Load BOM</h2>
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          className={`border-2 border-dashed rounded p-6 text-center transition-colors ${
            dragOver ? "border-black bg-gray-50" : "border-gray-300"
          }`}
        >
          <p className="text-sm text-gray-600 mb-3">
            Drag &amp; drop a CSV or XLSX file here, or choose one below.
            <br />
            Required columns: <code className="bg-gray-100 px-1 rounded">part_number</code>,{" "}
            <code className="bg-gray-100 px-1 rounded">qty</code>.
          </p>
          <div className="flex items-center justify-center gap-3 flex-wrap">
            <label className="inline-block rounded bg-black text-white px-4 py-2 cursor-pointer text-sm">
              Choose file
              <input
                type="file"
                accept=".csv,.xlsx,.xls"
                className="hidden"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) void handleFile(f);
                }}
              />
            </label>
            <button
              type="button"
              onClick={loadSample}
              className="rounded border border-gray-400 px-4 py-2 text-sm hover:bg-gray-50"
            >
              Load sample BOM
            </button>
          </div>
        </div>

        {fileName && (
          <div className="text-sm text-gray-700">
            Loaded <b>{fileName}</b> — {bom.length} line{bom.length === 1 ? "" : "s"}
          </div>
        )}
      </section>

      {/* Run card */}
      <section className="border rounded-lg p-4 space-y-3">
        <div className="flex items-center justify-between flex-wrap gap-2">
          <h2 className="font-semibold">2. Optimize</h2>
          <button
            onClick={runOptimize}
            disabled={!hasBom || loading}
            className="rounded bg-black text-white px-4 py-2 text-sm disabled:opacity-40"
          >
            {loading ? "Optimizing..." : "Run sourcing optimizer"}
          </button>
        </div>
        {error && <div className="text-red-600 text-sm">{error}</div>}
      </section>

      {/* Results */}
      {result && (
        <section className="space-y-4">
          <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
            <div className="border rounded p-3">
              <div className="text-xs text-gray-500">Lines</div>
              <div className="text-xl font-bold">{result.summary.lines_total}</div>
            </div>
            <div className="border rounded p-3">
              <div className="text-xs text-gray-500">Fully sourced</div>
              <div className="text-xl font-bold text-green-700">
                {result.summary.fully_sourced}
              </div>
            </div>
            <div className="border rounded p-3">
              <div className="text-xs text-gray-500">Substituted</div>
              <div className="text-xl font-bold text-blue-700">
                {result.summary.substituted}
              </div>
            </div>
            <div className="border rounded p-3">
              <div className="text-xs text-gray-500">No source</div>
              <div className="text-xl font-bold text-red-700">
                {result.summary.no_source}
              </div>
            </div>
            <div className="border rounded p-3">
              <div className="text-xs text-gray-500">Est. total cost</div>
              <div className="text-xl font-bold">
                ${result.summary.total_cost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </div>
              <div className="text-xs text-gray-500 mt-1">
                Critical lead: {result.summary.critical_lead_days}d
              </div>
            </div>
          </div>

          <div className="flex items-center justify-between flex-wrap gap-2">
            <h2 className="font-semibold">3. Sourcing Plan</h2>
            <button
              onClick={downloadPlanCsv}
              className="rounded border border-gray-400 px-3 py-1.5 text-sm hover:bg-gray-50"
            >
              ⬇ Download plan.csv
            </button>
          </div>

          {substitutions.length > 0 && (
            <div className="border rounded p-3 bg-blue-50 border-blue-200 text-sm">
              <div className="font-semibold mb-1">🔁 Substitutions applied ({substitutions.length})</div>
              <ul className="space-y-0.5">
                {substitutions.map((r, i) => (
                  <li key={i}>
                    <b>{r.substituted_from}</b> → <b>{r.picked_pn}</b> (qty {r.qty}) —{" "}
                    <span className="text-gray-700">{r.sub_reason}</span>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {unsourced.length > 0 && (
            <div
              className={`border rounded p-3 text-sm ${
                unsourcedSeverity === "alert"
                  ? "bg-red-50 border-red-200"
                  : "bg-amber-50 border-amber-200"
              }`}
            >
              <div className="font-semibold mb-1">
                {unsourcedSeverity === "alert" ? "⚠️" : "🟡"} Action required ({unsourced.length})
                {unsourcedSeverity === "warn" && (
                  <span className="ml-2 font-normal text-amber-800">
                    minor — most of the BOM is sourced
                  </span>
                )}
              </div>
              <ul className="space-y-0.5">
                {unsourced.map((r, i) => (
                  <li key={i}>
                    <b>{r.requested_pn}</b> (qty {r.qty}) — {r.note || "no source found"}
                  </li>
                ))}
              </ul>
            </div>
          )}

          <div className="border rounded overflow-x-auto">
            <table className="w-full text-sm">
              <caption className="sr-only">Sourcing plan rows</caption>
              <thead className="bg-gray-50 border-b">
                <tr className="text-left">
                  <th className="p-2"></th>
                  <th className="p-2">Part</th>
                  <th className="p-2 text-right">Qty</th>
                  <th className="p-2">Status</th>
                  <th className="p-2">Source</th>
                  <th className="p-2 text-right">Unit $</th>
                  <th className="p-2 text-right">Ext $</th>
                  <th className="p-2 text-right">Lead</th>
                </tr>
              </thead>
              <tbody>
                {result.rows.map((r, i) => {
                  const isExpanded = expanded.has(i);
                  const canExpand = r.alternates.length > 0 || r.tried_parts.length > 1;
                  return (
                    <tr
                      key={i}
                      className={`border-b ${
                        r.status === "no_source" ? "bg-red-50" : ""
                      } ${r.substituted_from ? "bg-blue-50" : ""}`}
                    >
                      <td className="p-2 align-top">
                        {canExpand && (
                          <button
                            onClick={() => toggleExpand(i)}
                            className="text-xs text-gray-600 hover:text-black"
                            aria-label="Expand row"
                          >
                            {isExpanded ? "▼" : "▶"}
                          </button>
                        )}
                      </td>
                      <td className="p-2 align-top">
                        <div className="font-mono font-semibold">{r.picked_pn}</div>
                        {r.substituted_from && (
                          <div className="text-xs text-blue-700">
                            🔁 from <span className="font-mono">{r.substituted_from}</span>
                          </div>
                        )}
                        {isExpanded && (
                          <div className="mt-2 space-y-1 text-xs text-gray-700">
                            {r.tried_parts.length > 1 && (
                              <div>
                                <span className="font-semibold">Tried: </span>
                                {r.tried_parts.join(" → ")}
                              </div>
                            )}
                            {r.alternates.length > 0 && (
                              <div>
                                <div className="font-semibold mt-1">Other sources:</div>
                                <ul className="ml-2">
                                  {r.alternates.map((a, j) => {
                                    const effQty = Math.max(r.qty, a.moq || 1);
                                    const ext = a.unit_price * effQty;
                                    return (
                                      <li key={j} className="font-mono">
                                        {a.source} — ${a.unit_price.toFixed(4)} × {effQty} = $
                                        {ext.toFixed(2)} ({a.lead_time_days}d, stock {a.stock})
                                      </li>
                                    );
                                  })}
                                </ul>
                              </div>
                            )}
                            {r.note && (
                              <div className="italic text-gray-600">note: {r.note}</div>
                            )}
                          </div>
                        )}
                      </td>
                      <td className="p-2 align-top text-right tabular-nums">{r.qty}</td>
                      <td className="p-2 align-top">
                        {r.status === "ok" && (
                          <span className="text-xs px-2 py-0.5 rounded bg-green-100 text-green-800">
                            ok
                          </span>
                        )}
                        {r.status === "partial_stock" && (
                          <span className="text-xs px-2 py-0.5 rounded bg-yellow-100 text-yellow-800">
                            partial
                          </span>
                        )}
                        {r.status === "no_source" && (
                          <span className="text-xs px-2 py-0.5 rounded bg-red-100 text-red-800">
                            no source
                          </span>
                        )}
                      </td>
                      <td className="p-2 align-top font-mono">{r.best?.source ?? "—"}</td>
                      <td className="p-2 align-top text-right tabular-nums">
                        {r.best ? `$${r.best.unit_price.toFixed(4)}` : "—"}
                      </td>
                      <td className="p-2 align-top text-right tabular-nums font-semibold">
                        {r.best ? `$${r.ext_price.toFixed(2)}` : "—"}
                      </td>
                      <td className="p-2 align-top text-right tabular-nums">
                        {r.best ? `${r.best.lead_time_days}d` : "—"}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Buy List — per vendor cart files */}
          {buyList.length > 0 && (
            <div className="space-y-3">
              <div className="flex items-center justify-between flex-wrap gap-2">
                <div>
                  <h2 className="font-semibold">4. Buy List</h2>
                  <p className="text-xs text-gray-600">
                    Sourcing plan grouped by vendor. Download per-vendor cart CSVs ready to paste
                    into DigiKey / Mouser / Arrow / LCSC bulk-add forms. MPN-based, works in mock
                    mode.
                  </p>
                </div>
                {buyList.length > 1 && (
                  <button
                    onClick={downloadAllVendorCsvs}
                    className="rounded bg-black text-white px-3 py-1.5 text-sm hover:bg-gray-800"
                  >
                    ⬇ Download all ({buyList.length} files)
                  </button>
                )}
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {buyList.map((bucket) => (
                  <div key={bucket.source} className="border rounded p-3 space-y-2">
                    <div className="flex items-center justify-between gap-2">
                      <div>
                        <div className="font-semibold text-base">{bucket.source}</div>
                        <div className="text-xs text-gray-500">
                          {bucket.line_count} line{bucket.line_count === 1 ? "" : "s"} · longest lead{" "}
                          {bucket.longest_lead}d
                        </div>
                      </div>
                      <div className="text-right">
                        <div className="text-lg font-bold tabular-nums">
                          ${bucket.total.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </div>
                      </div>
                    </div>
                    <details className="text-xs">
                      <summary className="cursor-pointer text-gray-600 hover:text-black">
                        Show parts
                      </summary>
                      <ul className="mt-2 space-y-0.5 font-mono">
                        {bucket.lines.map((l, i) => (
                          <li key={i} className="flex justify-between gap-2">
                            <span className="truncate">
                              {l.picked_pn}
                              {l.substituted_from && (
                                <span className="ml-1 text-blue-700">(🔁 {l.substituted_from})</span>
                              )}
                            </span>
                            <span className="tabular-nums whitespace-nowrap">
                              {l.qty} × ${l.unit_price.toFixed(4)} = ${l.ext_price.toFixed(2)}
                            </span>
                          </li>
                        ))}
                      </ul>
                    </details>
                    <button
                      onClick={() => downloadVendorCsv(bucket)}
                      className="w-full rounded border border-gray-400 px-3 py-1.5 text-sm hover:bg-gray-50"
                    >
                      ⬇ Download {bucket.source} cart
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </section>
      )}

      {!result && hasBom && !loading && !error && (
        <div className="text-sm text-gray-600">
          {bom.length} line{bom.length === 1 ? "" : "s"} loaded. Click <b>Run sourcing optimizer</b> above.
        </div>
      )}
    </main>
  );
}
