"use client";

import { useEffect, useState } from "react";

type ApiError = { message?: string };

type BomItem = { assembly_pn: string; revision: string; line_count: number };
type FloorStockItem = { component_pn: string; location: string | null };
type MasterBomSummary = {
  total_csv_rows: number;
  skipped_item_type: number;
  skipped_phantom: number;
  skipped_zero_qty: number;
  parsed_lines: number;
  consolidated_lines: number;
  duplicates_merged: number;
  unique_assemblies: number;
  pwb_assemblies: number;
  harness_assemblies: number;
  unique_components: number;
};
type MasterBomPreview = {
  summary: MasterBomSummary;
  sample_assemblies: Array<{ assembly_pn: string; revision: string; line_count: number; bom_type?: string }>;
  assemblies: Array<{ assembly_pn: string; revision: string; active: boolean }>;
  lines: Array<{ assembly_pn: string; revision: string; component_pn: string; qty_per_board: number }>;
};

function parseMasterCsvRow(line: string): string[] {
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
}

function extractRevision(revisionName: string, bomName: string): string {
  const prefixStripped = revisionName.startsWith(bomName)
    ? revisionName.slice(bomName.length).trim()
    : revisionName.trim();
  const revMatch = prefixStripped.match(/^Rev[_\s]*(.*)$/i);
  if (revMatch) {
    const rev = revMatch[1].trim();
    return rev.split(/\s/)[0] || rev || "1";
  }
  return prefixStripped || "1";
}

function parseMasterBomCsv(
  text: string,
  opts: { includeNonInvt: boolean; includePhantom: boolean }
): MasterBomPreview {
  const rawRows = text.split(/\r?\n/).filter(Boolean);
  if (rawRows.length < 2) throw new Error("CSV too short");
  const header = parseMasterCsvRow(rawRows[0]).map((h) => h.toLowerCase().replace(/\s+/g, "_"));
  const idx = {
    bom_name: header.indexOf("bom_name"),
    revision_name: header.indexOf("revision_name"),
    item: header.indexOf("item"),
    item_type: header.indexOf("item_type"),
    description: header.indexOf("description"),
    quantity: header.indexOf("quantity"),
    item_source: header.indexOf("item_source"),
  };
  if (idx.bom_name < 0 || idx.item < 0 || idx.quantity < 0) {
    throw new Error("Missing required columns: BOM Name, Item, Quantity");
  }

  // Pre-scan: classify each BOM as PWB or HARNESS before filtering
  // Rule 1: 6F- Assembly with PWB/PWA in description
  // Rule 2: 1C- item whose middle number matches BOM middle number (schematic)
  const pwbBoms = new Set<string>();
  for (let i = 1; i < rawRows.length; i++) {
    const r = parseMasterCsvRow(rawRows[i]);
    const bomName = (r[idx.bom_name] || "").trim();
    const item = (r[idx.item] || "").trim();
    const itemType = (r[idx.item_type] || "").trim();
    const desc = (r[idx.description] || "").trim().toUpperCase();
    if (!bomName) continue;
    // Rule 1: 6F with PWB/PWA
    if (item.startsWith("6F-") && itemType === "Assembly" && (desc.includes("PWB") || desc.includes("PWA"))) {
      pwbBoms.add(bomName);
    }
    // Rule 2: 1C schematic matching BOM middle number
    if (bomName.startsWith("8E-") && item.startsWith("1C-")) {
      const bomParts = bomName.split("-");
      const itemParts = item.split("-");
      if (bomParts.length >= 2 && itemParts.length >= 2 && bomParts[1] === itemParts[1]) {
        pwbBoms.add(bomName);
      }
    }
  }

  let skippedType = 0, skippedSource = 0, skippedZeroQty = 0;
  type Parsed = { assembly_pn: string; revision: string; component_pn: string; qty_per_board: number };
  const parsed: Parsed[] = [];
  for (let i = 1; i < rawRows.length; i++) {
    const r = parseMasterCsvRow(rawRows[i]);
    const bomName = (r[idx.bom_name] || "").trim();
    const revName = (r[idx.revision_name] || "").trim();
    const item = (r[idx.item] || "").trim();
    const itemType = (r[idx.item_type] || "").trim();
    const itemSource = (r[idx.item_source] || "").trim();
    const qty = Number(r[idx.quantity] || 0);
    if (!bomName || !item) continue;
    // Skip non-physical items: schematics, labels, sealants/coatings, paints/inks
    if (item.startsWith("1C-") || item.startsWith("0W-") || item.startsWith("0V-") || item.startsWith("0F-")) { skippedType++; continue; }
    if (!opts.includeNonInvt && itemType === "NonInvtPart") { skippedType++; continue; }
    if (itemType === "OthCharge") { skippedType++; continue; }
    if (!opts.includePhantom && itemSource === "PHANTOM") { skippedSource++; continue; }
    if (qty <= 0.01) { skippedZeroQty++; continue; }
    parsed.push({ assembly_pn: bomName, revision: extractRevision(revName, bomName), component_pn: item, qty_per_board: qty });
  }
  const cMap = new Map<string, Parsed>();
  for (const p of parsed) {
    const key = `${p.assembly_pn}||${p.revision}||${p.component_pn}`;
    const ex = cMap.get(key);
    if (ex) ex.qty_per_board += p.qty_per_board;
    else cMap.set(key, { ...p });
  }
  const lines = Array.from(cMap.values());
  const asmMap = new Map<string, { revision: string; count: number }>();
  for (const l of lines) {
    const key = `${l.assembly_pn}||${l.revision}`;
    const ex = asmMap.get(key);
    if (ex) ex.count++;
    else asmMap.set(key, { revision: l.revision, count: 1 });
  }
  const assemblies = Array.from(asmMap.entries()).map(([key, val]) => {
    const asm = key.split("||")[0];
    return {
      assembly_pn: asm, revision: val.revision, active: true,
      bom_type: pwbBoms.has(asm) ? "PWB" : "HARNESS",
    };
  });
  const pwbCount = assemblies.filter((a) => a.bom_type === "PWB").length;
  const harnessCount = assemblies.filter((a) => a.bom_type === "HARNESS").length;
  const sample = Array.from(asmMap.entries()).slice(0, 20).map(([key, val]) => {
    const asm = key.split("||")[0];
    return {
      assembly_pn: asm, revision: val.revision, line_count: val.count,
      bom_type: pwbBoms.has(asm) ? "PWB" : "HARNESS",
    };
  });
  return {
    summary: {
      total_csv_rows: rawRows.length - 1,
      skipped_item_type: skippedType,
      skipped_phantom: skippedSource,
      skipped_zero_qty: skippedZeroQty,
      parsed_lines: parsed.length,
      consolidated_lines: lines.length,
      duplicates_merged: parsed.length - lines.length,
      unique_assemblies: asmMap.size,
      pwb_assemblies: pwbCount,
      harness_assemblies: harnessCount,
      unique_components: new Set(lines.map((l) => l.component_pn)).size,
    },
    sample_assemblies: sample,
    assemblies,
    lines,
  };
}


export default function DataPage() {
  const [assembly, setAssembly] = useState("8E-03918-92");
  const [revision, setRevision] = useState("A");
  const [importMsg, setImportMsg] = useState<string | null>(null);
  const [boms, setBoms] = useState<BomItem[]>([]);

  const [floorStock, setFloorStock] = useState<FloorStockItem[]>([]);
  const [openPOCount, setOpenPOCount] = useState<number>(0);
  const [smtPartCount, setSmtPartCount] = useState<number>(0);

  const [adminToken, setAdminToken] = useState("");
  const [requiresAuth, setRequiresAuth] = useState(false);
  const [authed, setAuthed] = useState(false);

  const [masterBomPreview, setMasterBomPreview] = useState<MasterBomPreview | null>(null);
  const [masterBomFile, setMasterBomFile] = useState<File | null>(null);
  const [masterBomImporting, setMasterBomImporting] = useState(false);
  const [includeNonInvt, setIncludeNonInvt] = useState(false);
  const [includePhantom, setIncludePhantom] = useState(false);
  const [masterBomDbInfo, setMasterBomDbInfo] = useState<{ assemblies: number; bom_lines: number } | null>(null);

  async function loadBoms() {
    const res = await fetch("/api/boms");
    const data = await res.json();
    setBoms(data.items || []);
  }

  async function loadFloorStock() {
    const res = await fetch("/api/import/floor-stock");
    const data = await res.json();
    setFloorStock(data.items || []);
  }

  async function loadOpenPOCount() {
    const res = await fetch("/api/import/open-po");
    const data = await res.json();
    setOpenPOCount((data.items || []).length);
  }

  async function loadSmtPartCount() {
    try {
      const res = await fetch("/api/import/smt-parts");
      const data = await res.json();
      setSmtPartCount(data.count || 0);
    } catch { /* silent */ }
  }

  async function importSmtParts(file: File) {
    setImportMsg("Importing SMT parts list...");
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/api/import/smt-parts", { method: "POST", body: fd, headers: adminHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "SMT parts import failed");
    setImportMsg(`SMT parts imported: ${data.imported} parts`);
    await loadSmtPartCount();
  }

  async function refreshAdminStatus() {
    const res = await fetch("/api/admin/status", { cache: "no-store" });
    const data = await res.json();
    setRequiresAuth(!!data.requiresAuth);
    setAuthed(!!data.authed);
  }

  async function loadMasterBomInfo() {
    try {
      const res = await fetch("/api/import/master-bom", { headers: adminHeaders() });
      if (res.ok) {
        const data = await res.json();
        setMasterBomDbInfo(data);
      }
    } catch { /* silent */ }
  }

  async function previewMasterBom(file: File) {
    setMasterBomFile(file);
    setImportMsg("Parsing Master BOM in browser...");
    try {
      const text = await file.text();
      const preview = parseMasterBomCsv(text, { includeNonInvt, includePhantom });
      setMasterBomPreview(preview);
      setImportMsg(null);
    } catch (err: unknown) {
      setImportMsg((err as ApiError)?.message || "Failed to parse CSV");
    }
  }

  async function confirmMasterBomImport() {
    if (!masterBomPreview) return;
    const ok = window.confirm(
      `⚠️ This will REPLACE the entire BOM database with ${masterBomPreview.summary.unique_assemblies} assemblies and ${masterBomPreview.summary.consolidated_lines} BOM lines. Continue?`
    );
    if (!ok) return;
    setMasterBomImporting(true);
    try {
      setImportMsg("Clearing old BOM data...");
      const clearRes = await fetch("/api/import/master-bom", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...adminHeaders() },
        body: JSON.stringify({ action: "clear" }),
      });
      if (!clearRes.ok) { const d = await clearRes.json(); throw new Error(d?.error || "Clear failed"); }

      const BATCH = 400;
      const { assemblies, lines } = masterBomPreview;
      const totalBatches = Math.ceil(assemblies.length / BATCH) + Math.ceil(lines.length / BATCH);
      let batchNum = 0;

      for (let i = 0; i < assemblies.length; i += BATCH) {
        batchNum++;
        setImportMsg(`Uploading assemblies... batch ${batchNum}/${totalBatches}`);
        const batch = assemblies.slice(i, i + BATCH);
        const res = await fetch("/api/import/master-bom", {
          method: "POST",
          headers: { "Content-Type": "application/json", ...adminHeaders() },
          body: JSON.stringify({ action: "batch", assemblies: batch, lines: [] }),
        });
        if (!res.ok) { const d = await res.json(); throw new Error(d?.error || "Assembly batch failed"); }
      }

      for (let i = 0; i < lines.length; i += BATCH) {
        batchNum++;
        setImportMsg(`Uploading BOM lines... batch ${batchNum}/${totalBatches}`);
        const batch = lines.slice(i, i + BATCH);
        const res = await fetch("/api/import/master-bom", {
          method: "POST",
          headers: { "Content-Type": "application/json", ...adminHeaders() },
          body: JSON.stringify({ action: "batch", assemblies: [], lines: batch }),
        });
        if (!res.ok) { const d = await res.json(); throw new Error(d?.error || "BOM line batch failed"); }
      }

      setImportMsg("Finalizing...");
      const finRes = await fetch("/api/import/master-bom", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...adminHeaders() },
        body: JSON.stringify({ action: "finalize" }),
      });
      const finData = await finRes.json();
      setImportMsg(`✅ Imported ${finData.assemblies} assemblies (${finData.pwb_assemblies || 0} PWB, ${finData.harness_assemblies || 0} Harness) with ${finData.bom_lines?.toLocaleString()} BOM lines`);
      setMasterBomPreview(null);
      setMasterBomFile(null);
      await loadBoms();
      await loadMasterBomInfo();
    } catch (err: unknown) {
      setImportMsg((err as ApiError)?.message || "Master BOM import failed");
    } finally {
      setMasterBomImporting(false);
    }
  }

  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect
    void loadBoms();
    void loadFloorStock();
    void loadOpenPOCount();
    void loadSmtPartCount();
    void refreshAdminStatus();
    void loadMasterBomInfo();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function adminHeaders(): Record<string, string> {
    const h: Record<string, string> = {};
    if (adminToken) h["x-admin-token"] = adminToken;
    return h;
  }

  async function loginAdmin() {
    const res = await fetch("/api/admin/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token: adminToken }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Login failed");

    setImportMsg(data?.authDisabled ? "Admin token not configured. Write actions are open." : "Admin login successful.");
    await refreshAdminStatus();
  }

  async function logoutAdmin() {
    await fetch("/api/admin/logout", { method: "POST" });
    setAuthed(false);
    setAdminToken("");
    setImportMsg("Admin session cleared.");
  }

  async function importInventory(file: File) {
    setImportMsg("Importing inventory...");
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/api/import/inventory", { method: "POST", body: fd, headers: adminHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Inventory import failed");
    setImportMsg(`Inventory imported: ${data.imported}`);
  }

  async function importFloorStock(file: File) {
    setImportMsg("Importing floor stock...");
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/api/import/floor-stock", { method: "POST", body: fd, headers: adminHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Floor stock import failed");
    setImportMsg(`Floor stock imported: ${data.imported} parts`);
    await loadFloorStock();
  }

  async function deleteFloorStockPart(pn: string) {
    const res = await fetch("/api/import/floor-stock", {
      method: "DELETE",
      headers: { "Content-Type": "application/json", ...adminHeaders() },
      body: JSON.stringify({ component_pn: pn }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Delete failed");
    setImportMsg(`Removed floor stock part: ${pn}`);
    await loadFloorStock();
  }

  async function importOpenPO(file: File) {
    setImportMsg("Importing open purchase orders...");
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/api/import/open-po", { method: "POST", body: fd, headers: adminHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Open PO import failed");
    setImportMsg(`Open POs imported: ${data.imported} lines`);
    await loadOpenPOCount();
  }

  async function importStockroom(file: File) {
    setImportMsg("Importing stockroom...");
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch("/api/import/stockroom", { method: "POST", body: fd, headers: adminHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Stockroom import failed");
    setImportMsg(`Stockroom imported: ${data.imported}`);
  }

  async function importBom(file: File) {
    const bomAssembly = (window.prompt("BOM Assembly Part Number?", assembly) || "").trim();
    if (!bomAssembly) throw new Error("BOM Assembly Part Number is required");

    const bomRevision = (window.prompt("BOM Revision?", revision) || "").trim();
    if (!bomRevision) throw new Error("BOM Revision is required");

    const preview = new FormData();
    preview.append("assembly_pn", bomAssembly);
    preview.append("revision", bomRevision);
    preview.append("file", file);
    preview.append("dry_run", "1");
    const pRes = await fetch("/api/import/bom", { method: "POST", body: preview, headers: adminHeaders() });
    const pData = await pRes.json();
    if (!pRes.ok) throw new Error(pData?.error || "BOM precheck failed");

    const rc = pData?.revisionControl;
    if (rc?.oldRevsToRemove?.length > 0) {
      const ok = window.confirm(
        `Revision upgrade: Importing rev ${rc.incomingRev} will remove older revision(s): ${rc.oldRevsToRemove.join(", ")}. Continue?`
      );
      if (!ok) {
        setImportMsg("BOM import canceled.");
        return;
      }
    } else if ((pData?.existingCount || 0) > 0) {
      const ok = window.confirm(
        `Rev overwrite warning: ${bomAssembly} rev ${bomRevision} already has ${pData.existingCount} lines. Replace with ${pData.incomingCount} new lines?`
      );
      if (!ok) {
        setImportMsg("BOM import canceled.");
        return;
      }
    }

    setImportMsg("Importing BOM...");
    const fd = new FormData();
    fd.append("assembly_pn", bomAssembly);
    fd.append("revision", bomRevision);
    fd.append("file", file);
    const res = await fetch("/api/import/bom", { method: "POST", body: fd, headers: adminHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "BOM import failed");

    setAssembly(bomAssembly);
    setRevision(bomRevision);
    const merged = Number(data?.duplicateRowsMerged || 0);
    const removedRevs = data?.oldRevsRemoved;
    const verified = data?.verified ?? "?";
    setImportMsg(
      `BOM imported: ${data.imported} lines for ${bomAssembly} rev ${bomRevision} (verified: ${verified})${merged > 0 ? ` (merged ${merged} duplicate rows)` : ""}${removedRevs ? ` | Old rev(s) removed: ${removedRevs.join(", ")}` : ""}`
    );
    await loadBoms();
  }

  async function deleteBom(targetAssembly?: string, targetRev?: string) {
    const a = (targetAssembly || window.prompt("Delete BOM Assembly Part Number?", assembly) || "").trim();
    if (!a) return;
    const r = (targetRev || window.prompt("Delete BOM Revision?", revision) || "").trim();
    if (!r) return;

    const ok = window.confirm(`Delete BOM ${a} rev ${r}? This cannot be undone from app.`);
    if (!ok) return;

    const res = await fetch("/api/import/bom/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json", ...adminHeaders() },
      body: JSON.stringify({ assembly_pn: a, revision: r }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Delete BOM failed");
    setImportMsg(`Deleted BOM ${a} rev ${r}: ${data.deleted} lines removed.`);
    await loadBoms();
  }

  const canWrite = !requiresAuth || authed;

  return (
    <main className="max-w-4xl mx-auto p-6 space-y-6 w-full" onDragOver={(e) => e.preventDefault()} onDrop={(e) => e.preventDefault()}>
      <h1 className="text-2xl font-bold">Data Manager</h1>

      <section className="border rounded p-3 space-y-2">
        <h2 className="font-semibold">Admin Access</h2>
        {!requiresAuth ? (
          <p className="text-sm text-gray-600">Admin token not configured on server. Write actions are currently open.</p>
        ) : (
          <>
            <div className="flex gap-2">
              <input
                className="border rounded p-2 flex-1"
                type="password"
                placeholder="Enter admin token"
                value={adminToken}
                onChange={(e) => setAdminToken(e.target.value)}
              />
              <button
                className="rounded border px-3"
                onClick={async () => {
                  try {
                    await loginAdmin();
                  } catch (err: unknown) {
                    const e = err as ApiError;
                    setImportMsg(e?.message || "Admin login failed");
                  }
                }}
              >
                Login
              </button>
              <button className="rounded border px-3" onClick={logoutAdmin}>Logout</button>
            </div>
            <p className="text-sm text-gray-600">Status: {authed ? "Authenticated" : "Not authenticated"}</p>
          </>
        )}
      </section>

      <section className="border-2 border-indigo-200 rounded p-4 space-y-3 bg-indigo-50/30">
        <h2 className="font-semibold text-lg">📦 Master BOM Import (NetSuite Full Export)</h2>
        <p className="text-xs text-gray-600">Upload the full NetSuite BOM export CSV. This replaces all individual BOM uploads with one file.</p>
        {masterBomDbInfo && (
          <div className="text-xs text-gray-500">
            Current database: {masterBomDbInfo.assemblies} assemblies, {masterBomDbInfo.bom_lines.toLocaleString()} BOM lines
          </div>
        )}
        <div className="flex gap-4 items-center flex-wrap">
          <label className="border-2 border-dashed rounded p-3 block cursor-pointer hover:border-indigo-400 transition-colors flex-1 min-w-[200px]"
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("border-indigo-400", "bg-indigo-50"); }}
            onDragLeave={(e) => { e.currentTarget.classList.remove("border-indigo-400", "bg-indigo-50"); }}
            onDrop={async (e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("border-indigo-400", "bg-indigo-50");
              const f = e.dataTransfer.files[0];
              if (!f || !canWrite) return;
              try { await previewMasterBom(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Preview failed"); }
            }}
          >
            <div className="text-sm mb-1 font-medium">Drop Master BOM CSV here</div>
            <div className="text-xs text-gray-400">Expected columns: BOM Name, Revision Name, Item, Quantity, Item Type, Item Source</div>
            <input type="file" accept=".csv" disabled={!canWrite} onChange={async (e) => {
              const f = e.target.files?.[0];
              if (!f) return;
              try { await previewMasterBom(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Preview failed"); }
              e.target.value = "";
            }} />
          </label>
          <div className="flex flex-col gap-1 text-sm">
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={includeNonInvt} onChange={(e) => setIncludeNonInvt(e.target.checked)} />
              Include NonInvtPart
            </label>
            <label className="flex items-center gap-2">
              <input type="checkbox" checked={includePhantom} onChange={(e) => setIncludePhantom(e.target.checked)} />
              Include PHANTOM items
            </label>
          </div>
        </div>

        {masterBomPreview && (
          <div className="border rounded p-3 bg-white space-y-2">
            <h3 className="font-semibold text-sm">Preview Summary</h3>
            <div className="grid grid-cols-3 gap-2 text-sm">
              <div>CSV rows: <strong>{masterBomPreview.summary.total_csv_rows.toLocaleString()}</strong></div>
              <div>Unique assemblies: <strong>{masterBomPreview.summary.unique_assemblies.toLocaleString()}</strong> <span className="text-xs">(<span className="text-blue-700">{masterBomPreview.summary.pwb_assemblies} PWB</span> · <span className="text-amber-700">{masterBomPreview.summary.harness_assemblies} Harness</span>)</span></div>
              <div>Unique components: <strong>{masterBomPreview.summary.unique_components.toLocaleString()}</strong></div>
              <div>BOM lines to import: <strong>{masterBomPreview.summary.consolidated_lines.toLocaleString()}</strong></div>
              <div>Skipped (item type): <span className="text-gray-500">{masterBomPreview.summary.skipped_item_type.toLocaleString()}</span></div>
              <div>Skipped (phantom): <span className="text-gray-500">{masterBomPreview.summary.skipped_phantom.toLocaleString()}</span></div>
              <div>Skipped (zero qty): <span className="text-gray-500">{masterBomPreview.summary.skipped_zero_qty.toLocaleString()}</span></div>
              <div>Duplicates merged: <span className="text-gray-500">{masterBomPreview.summary.duplicates_merged.toLocaleString()}</span></div>
            </div>
            <div className="text-xs text-gray-500 mt-2">Sample assemblies:</div>
            <div className="text-xs font-mono max-h-32 overflow-auto">
              {masterBomPreview.sample_assemblies.map((a) => (
                <div key={`${a.assembly_pn}-${a.revision}`}>
                  <span className={`text-xs px-1 py-0.5 rounded mr-1 ${a.bom_type === "PWB" ? "bg-blue-100 text-blue-800" : "bg-amber-100 text-amber-800"}`}>{a.bom_type || "?"}</span>
                  {a.assembly_pn} rev {a.revision} ({a.line_count} lines)
                </div>
              ))}
            </div>
            <div className="flex gap-2 mt-2">
              <button
                className="bg-indigo-600 text-white px-4 py-2 rounded hover:bg-indigo-700 disabled:opacity-50"
                disabled={masterBomImporting}
                onClick={confirmMasterBomImport}
              >
                {masterBomImporting ? "Importing..." : `Import ${masterBomPreview.summary.unique_assemblies} Assemblies`}
              </button>
              <button className="border rounded px-4 py-2" onClick={() => { setMasterBomPreview(null); setMasterBomFile(null); }}>Cancel</button>
            </div>
          </div>
        )}
      </section>

      <section className="border rounded p-3 space-y-3">
        <h2 className="font-semibold">Individual Imports</h2>
        {!canWrite && <div className="text-sm text-red-700">Admin login required for write actions.</div>}
        <div className="grid md:grid-cols-6 gap-4">
          <label
            className="border-2 border-dashed rounded p-3 block cursor-pointer hover:border-gray-400 transition-colors"
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("border-blue-400", "bg-blue-50"); }}
            onDragLeave={(e) => { e.currentTarget.classList.remove("border-blue-400", "bg-blue-50"); }}
            onDrop={async (e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("border-blue-400", "bg-blue-50");
              const f = e.dataTransfer.files[0];
              if (!f || !canWrite) return;
              try { await importInventory(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Inventory import failed"); }
            }}
          >
            <div className="text-sm mb-2">Import XRAY Inventory CSV</div>
            <div className="text-xs text-gray-400 mb-2">Drop CSV here or click to browse</div>
            <input
              type="file"
              accept=".csv"
              disabled={!canWrite}
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                try { await importInventory(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Inventory import failed"); }
                e.target.value = "";
              }}
            />
          </label>

          <label
            className="border-2 border-dashed rounded p-3 block cursor-pointer hover:border-gray-400 transition-colors"
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("border-blue-400", "bg-blue-50"); }}
            onDragLeave={(e) => { e.currentTarget.classList.remove("border-blue-400", "bg-blue-50"); }}
            onDrop={async (e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("border-blue-400", "bg-blue-50");
              const f = e.dataTransfer.files[0];
              if (!f || !canWrite) return;
              try { await importStockroom(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Stockroom import failed"); }
            }}
          >
            <div className="text-sm mb-2">Import Stockroom CSV</div>
            <div className="text-xs text-gray-400 mb-2">Drop CSV here or click to browse</div>
            <input
              type="file"
              accept=".csv"
              disabled={!canWrite}
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                try { await importStockroom(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Stockroom import failed"); }
                e.target.value = "";
              }}
            />
          </label>

          <label
            className="border-2 border-dashed rounded p-3 block cursor-pointer hover:border-gray-400 transition-colors"
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("border-blue-400", "bg-blue-50"); }}
            onDragLeave={(e) => { e.currentTarget.classList.remove("border-blue-400", "bg-blue-50"); }}
            onDrop={async (e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("border-blue-400", "bg-blue-50");
              const f = e.dataTransfer.files[0];
              if (!f || !canWrite) return;
              try { await importBom(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "BOM import failed"); }
            }}
          >
            <div className="text-sm mb-2">Import BOM CSV (component_pn, qty_per_board)</div>
            <div className="text-xs text-gray-400 mb-2">Drop CSV here or click to browse</div>
            <input
              type="file"
              accept=".csv"
              disabled={!canWrite}
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                try { await importBom(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "BOM import failed"); }
                e.target.value = "";
              }}
            />
          </label>

          <label
            className="border-2 border-dashed rounded p-3 block cursor-pointer hover:border-gray-400 transition-colors"
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("border-blue-400", "bg-blue-50"); }}
            onDragLeave={(e) => { e.currentTarget.classList.remove("border-blue-400", "bg-blue-50"); }}
            onDrop={async (e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("border-blue-400", "bg-blue-50");
              const f = e.dataTransfer.files[0];
              if (!f || !canWrite) return;
              try { await importFloorStock(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Floor stock import failed"); }
            }}
          >
            <div className="text-sm mb-2">Import Floor Stock CSV (Part#, Location)</div>
            <div className="text-xs text-gray-400 mb-2">Drop CSV here or click to browse</div>
            <input
              type="file"
              accept=".csv"
              disabled={!canWrite}
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                try {
                  await importFloorStock(f);
                } catch (err: unknown) {
                  const e = err as ApiError;
                  setImportMsg(e?.message || "Floor stock import failed");
                }
                e.target.value = "";
              }}
            />
          </label>

          <label
            className="border-2 border-dashed rounded p-3 block cursor-pointer hover:border-gray-400 transition-colors"
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("border-blue-400", "bg-blue-50"); }}
            onDragLeave={(e) => { e.currentTarget.classList.remove("border-blue-400", "bg-blue-50"); }}
            onDrop={async (e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("border-blue-400", "bg-blue-50");
              const f = e.dataTransfer.files[0];
              if (!f || !canWrite) return;
              try { await importOpenPO(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Open PO import failed"); }
            }}
          >
            <div className="text-sm mb-2">Import Open Purchase Orders CSV (Optional)</div>
            <div className="text-xs text-gray-400 mb-1">Drop CSV here or click to browse</div>
            <div className="text-xs text-gray-400 mb-2">Data persists until new file is uploaded.</div>
            {openPOCount > 0 && <div className="text-xs text-green-600 mb-1">✅ {openPOCount} PO lines stored</div>}
            <input
              type="file"
              accept=".csv"
              disabled={!canWrite}
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                try { await importOpenPO(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "Open PO import failed"); }
                e.target.value = "";
              }}
            />
          </label>

          <label
            className="border-2 border-dashed rounded p-3 block cursor-pointer hover:border-gray-400 transition-colors"
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.classList.add("border-orange-400", "bg-orange-50"); }}
            onDragLeave={(e) => { e.currentTarget.classList.remove("border-orange-400", "bg-orange-50"); }}
            onDrop={async (e) => {
              e.preventDefault();
              e.currentTarget.classList.remove("border-orange-400", "bg-orange-50");
              const f = e.dataTransfer.files[0];
              if (!f || !canWrite) return;
              try { await importSmtParts(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "SMT parts import failed"); }
            }}
          >
            <div className="text-sm mb-2">Import SMT Parts List</div>
            <div className="text-xs text-gray-400 mb-1">Drop CSV here or click to browse</div>
            <div className="text-xs text-gray-400 mb-2">Tags BOM parts as SMT vs Through Hole.</div>
            {smtPartCount > 0 && <div className="text-xs text-green-600 mb-1">✅ {smtPartCount} SMT parts stored</div>}
            <input
              type="file"
              accept=".csv"
              disabled={!canWrite}
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                try { await importSmtParts(f); } catch (err: unknown) { setImportMsg((err as ApiError)?.message || "SMT parts import failed"); }
                e.target.value = "";
              }}
            />
          </label>
        </div>
        {importMsg && <div className="text-sm text-blue-700">{importMsg}</div>}
      </section>

      <section className="border rounded p-3 space-y-2">
        <h2 className="font-semibold">Imported BOM Library</h2>
        {boms.length === 0 ? (
          <p className="text-sm text-gray-600">No BOMs imported yet.</p>
        ) : (
          <ul className="text-sm space-y-1">
            {boms.map((b) => (
              <li key={`${b.assembly_pn}-${b.revision}`} className="flex items-center justify-between gap-2 border rounded p-2">
                <span>{b.assembly_pn} rev {b.revision} ({b.line_count} lines)</span>
                <button
                  className="rounded border px-2 py-1"
                  disabled={!canWrite}
                  onClick={async () => {
                    try {
                      await deleteBom(b.assembly_pn, b.revision);
                    } catch (err: unknown) {
                      const e = err as ApiError;
                      setImportMsg(e?.message || "Delete BOM failed");
                    }
                  }}
                >
                  Delete
                </button>
              </li>
            ))}
          </ul>
        )}
      </section>

      <section className="border rounded p-3 space-y-2">
        <h2 className="font-semibold">Floor Stock (FI) Parts ({floorStock.length})</h2>
        <p className="text-xs text-gray-500">Parts marked as floor stock are excluded from shortage counts in Build Readiness checks.</p>
        {floorStock.length === 0 ? (
          <p className="text-sm text-gray-600">No floor stock parts imported yet.</p>
        ) : (
          <div className="overflow-auto max-h-64">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left border-b">
                  <th className="p-2">Part Number</th>
                  <th className="p-2">Location</th>
                  <th className="p-2">Action</th>
                </tr>
              </thead>
              <tbody>
                {floorStock.map((f) => (
                  <tr key={f.component_pn} className="border-b">
                    <td className="p-2 font-mono">{f.component_pn}</td>
                    <td className="p-2">{f.location || "—"}</td>
                    <td className="p-2">
                      <button
                        className="rounded border px-2 py-1 text-xs"
                        disabled={!canWrite}
                        onClick={async () => {
                          try {
                            await deleteFloorStockPart(f.component_pn);
                          } catch (err: unknown) {
                            const e = err as ApiError;
                            setImportMsg(e?.message || "Delete failed");
                          }
                        }}
                      >
                        Remove
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </main>
  );
}
