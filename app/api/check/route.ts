import { NextRequest, NextResponse } from "next/server";
import fs from "node:fs";
import path from "node:path";
import { supabase } from "@/lib/supabase";

function parseCsv(filePath: string): string[][] {
  const raw = fs.readFileSync(filePath, "utf8");
  return raw
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => line.split(","));
}

async function checkFromSupabase(assembly: string, revision: string, qty: number) {
  if (!supabase) throw new Error("Supabase not configured");

  // Paginate bom_lines (filtered by assembly/rev, usually <1000 but be safe)
  const bomLines: Array<{ component_pn: string; qty_per_board: number }> = [];
  {
    const ps = 1000;
    let from = 0;
    while (true) {
      const { data, error } = await supabase
        .from("bom_lines")
        .select("component_pn, qty_per_board")
        .eq("assembly_pn", assembly)
        .eq("revision", revision)
        .range(from, from + ps - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      bomLines.push(...data);
      if (data.length < ps) break;
      from += ps;
    }
  }
  if (bomLines.length === 0) throw new Error("No BOM found for assembly/revision");

  // Paginate inventory_lots
  const invRows: Array<{ component_pn: string; qty_on_hand: number; qty_reserved: number }> = [];
  {
    const ps = 1000;
    let from = 0;
    while (true) {
      const { data, error } = await supabase
        .from("inventory_lots")
        .select("component_pn, qty_on_hand, qty_reserved")
        .range(from, from + ps - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      invRows.push(...data);
      if (data.length < ps) break;
      from += ps;
    }
  }

  const invByPart = new Map<string, number>();
  for (const r of invRows || []) {
    const avail = Math.max(0, Number(r.qty_on_hand || 0) - Number(r.qty_reserved || 0));
    invByPart.set(r.component_pn, (invByPart.get(r.component_pn) || 0) + avail);
  }

  const shortages: Array<{ part: string; required: number; available: number; shortage: number }> = [];
  let maxBuildable = Number.MAX_SAFE_INTEGER;

  for (const l of bomLines) {
    const qpb = Number(l.qty_per_board || 0);
    const required = qpb * qty;
    const available = invByPart.get(l.component_pn) || 0;
    const shortage = Math.max(0, required - available);
    if (shortage > 0) shortages.push({ part: l.component_pn, required, available, shortage });
    if (qpb > 0) maxBuildable = Math.min(maxBuildable, Math.floor(available / qpb));
  }

  const canRun = shortages.length === 0;
  const finalMax = maxBuildable === Number.MAX_SAFE_INTEGER ? 0 : maxBuildable;

  // Save history row
  const { data: checkRow, error: checkErr } = await supabase
    .from("build_checks")
    .insert({
      assembly_pn: assembly,
      revision,
      build_qty: qty,
      can_run: canRun,
      shortage_count: shortages.length,
      near_empty_count: 0,
      max_buildable: finalMax,
      created_by: "vu",
    })
    .select("check_id")
    .single();

  if (!checkErr && checkRow?.check_id) {
    const lines = bomLines.map((l) => {
      const qpb = Number(l.qty_per_board || 0);
      const required = qpb * qty;
      const available = invByPart.get(l.component_pn) || 0;
      const shortage = Math.max(0, required - available);
      return {
        check_id: checkRow.check_id,
        component_pn: l.component_pn,
        qty_per_board: qpb,
        required_qty: required,
        available_qty: available,
        shortage_qty: shortage,
        near_empty_5pct: false,
      };
    });
    await supabase.from("build_check_lines").insert(lines);
  }

  return {
    canRun,
    maxBuildable: finalMax,
    shortages,
    source: "supabase",
    checkId: checkRow?.check_id ?? null,
  };
}

function checkFromCsv(assembly: string, revision: string, qty: number) {
  const root = process.cwd();
  const bomPath = path.join(root, "data", "bom.csv");
  const invPath = path.join(root, "data", "inventory_lots.csv");

  const bomRows = parseCsv(bomPath);
  const invRows = parseCsv(invPath);

  const invByPart = new Map<string, number>();
  for (let i = 1; i < invRows.length; i++) {
    const r = invRows[i];
    const part = (r[1] || "").trim();
    const onHand = Number(r[2] || 0);
    const reserved = Number(r[3] || 0);
    const avail = Math.max(0, onHand - reserved);
    invByPart.set(part, (invByPart.get(part) || 0) + avail);
  }

  const lines = [] as Array<{ part: string; qpb: number }>;
  for (let i = 1; i < bomRows.length; i++) {
    const r = bomRows[i];
    if ((r[0] || "").trim() === assembly && (r[1] || "").trim() === revision) {
      lines.push({ part: (r[2] || "").trim(), qpb: Number(r[3] || 0) });
    }
  }

  if (lines.length === 0) throw new Error("No BOM found for assembly/revision");

  const shortages: Array<{ part: string; required: number; available: number; shortage: number }> = [];
  let maxBuildable = Number.MAX_SAFE_INTEGER;

  for (const l of lines) {
    const required = l.qpb * Number(qty);
    const available = invByPart.get(l.part) || 0;
    const shortage = Math.max(0, required - available);
    if (shortage > 0) shortages.push({ part: l.part, required, available, shortage });
    if (l.qpb > 0) maxBuildable = Math.min(maxBuildable, Math.floor(available / l.qpb));
  }

  return {
    canRun: shortages.length === 0,
    maxBuildable: maxBuildable === Number.MAX_SAFE_INTEGER ? 0 : maxBuildable,
    shortages,
    source: "csv-fallback",
    checkId: null,
  };
}

export async function POST(req: NextRequest) {
  try {
    const { assembly, revision, qty } = await req.json();
    if (!assembly || !revision || !qty) {
      return NextResponse.json({ error: "assembly, revision, qty required" }, { status: 400 });
    }

    try {
      const result = await checkFromSupabase(assembly, revision, Number(qty));
      return NextResponse.json(result);
    } catch {
      const result = checkFromCsv(assembly, revision, Number(qty));
      return NextResponse.json(result);
    }
  } catch (e: any) {
    return NextResponse.json({ error: e.message || "Unexpected error" }, { status: 500 });
  }
}
