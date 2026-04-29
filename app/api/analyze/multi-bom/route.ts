import { NextRequest, NextResponse } from "next/server";
import fs from "node:fs";
import path from "node:path";
import { supabase } from "@/lib/supabase";

type PlanInput = {
  run_id?: string;
  run_label?: string;
  assembly_pn: string;
  revision?: string;
  qty: number;
  priority?: number;
};

type PlanComputed = {
  run_id: string;
  run_label: string;
  assembly_pn: string;
  revision: string;
  qty: number;
  priority: number;
  max_buildable: number;
  can_complete: boolean;
  shortages: Array<{ part: string; required: number; available: number; shortage: number }>;
  requirements: Map<string, number>;
};

function parseCsv(filePath: string): string[][] {
  const raw = fs.readFileSync(filePath, "utf8");
  return raw
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => line.split(",").map((x) => x.trim()));
}

function isFeasibleSubset(items: PlanComputed[], available: Map<string, number>): boolean {
  const required = new Map<string, number>();
  for (const it of items) {
    for (const [part, qty] of it.requirements.entries()) {
      required.set(part, (required.get(part) || 0) + qty);
      if ((required.get(part) || 0) > (available.get(part) || 0)) return false;
    }
  }
  return true;
}

function scoreSubset(items: PlanComputed[]) {
  return {
    completeRuns: items.length,
    boards: items.reduce((s, x) => s + x.qty, 0),
    priority: items.reduce((s, x) => s + x.priority, 0),
  };
}

function betterScore(a: ReturnType<typeof scoreSubset>, b: ReturnType<typeof scoreSubset>): boolean {
  if (a.completeRuns !== b.completeRuns) return a.completeRuns > b.completeRuns;
  if (a.priority !== b.priority) return a.priority > b.priority;
  return a.boards > b.boards;
}

function findBestSubset(candidates: PlanComputed[], available: Map<string, number>): PlanComputed[] {
  // Exact search for small sets
  if (candidates.length <= 18) {
    let best: PlanComputed[] = [];
    const n = candidates.length;

    function dfs(i: number, chosen: PlanComputed[]) {
      if (i === n) {
        if (!isFeasibleSubset(chosen, available)) return;
        if (betterScore(scoreSubset(chosen), scoreSubset(best))) best = [...chosen];
        return;
      }

      dfs(i + 1, chosen);
      chosen.push(candidates[i]);
      // lightweight early feasibility check
      if (isFeasibleSubset(chosen, available)) dfs(i + 1, chosen);
      chosen.pop();
    }

    dfs(0, []);
    return best;
  }

  // Greedy fallback for large sets
  const sorted = [...candidates].sort((a, b) => b.priority - a.priority || b.qty - a.qty);
  const chosen: PlanComputed[] = [];
  for (const c of sorted) {
    const next = [...chosen, c];
    if (isFeasibleSubset(next, available)) chosen.push(c);
  }
  return chosen;
}

async function analyzeFromSupabase(plans: PlanInput[]) {
  if (!supabase) throw new Error("Supabase not configured");

  // Paginate bom_lines to avoid Supabase 1000-row default limit
  const bomLines: Array<{ assembly_pn: string; revision: string; component_pn: string; qty_per_board: number }> = [];
  {
    const pageSize = 1000;
    let from = 0;
    while (true) {
      const { data, error } = await supabase
        .from("bom_lines")
        .select("assembly_pn, revision, component_pn, qty_per_board")
        .range(from, from + pageSize - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      bomLines.push(...data);
      if (data.length < pageSize) break;
      from += pageSize;
    }
  }

  // Paginate inventory_lots similarly
  const invRows: Array<{ component_pn: string; qty_on_hand: number; qty_reserved: number; status: string }> = [];
  {
    const pageSize = 1000;
    let from = 0;
    while (true) {
      const { data, error } = await supabase
        .from("inventory_lots")
        .select("component_pn, qty_on_hand, qty_reserved, status")
        .range(from, from + pageSize - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      invRows.push(...data);
      if (data.length < pageSize) break;
      from += pageSize;
    }
  }

  const available = new Map<string, number>();
  for (const r of invRows) {
    if (String(r.status || "ACTIVE") !== "ACTIVE") continue;
    const avail = Math.max(0, Number(r.qty_on_hand || 0) - Number(r.qty_reserved || 0));
    available.set(r.component_pn, (available.get(r.component_pn) || 0) + avail);
  }

  const byAssembly = new Map<string, Array<{ revision: string; component_pn: string; qty_per_board: number }>>();
  for (const r of bomLines) {
    const key = r.assembly_pn;
    if (!byAssembly.has(key)) byAssembly.set(key, []);
    byAssembly.get(key)!.push({ revision: r.revision, component_pn: r.component_pn, qty_per_board: Number(r.qty_per_board || 0) });
  }

  const computed: PlanComputed[] = [];
  const errors: string[] = [];

  for (const p of plans) {
    const assembly = String(p.assembly_pn || "").trim();
    const qty = Math.max(1, Math.floor(Number(p.qty || 0)));
    const priority = Number(p.priority || 1);
    if (!assembly || qty <= 0) continue;

    const rows = byAssembly.get(assembly) || [];
    if (rows.length === 0) {
      errors.push(`No BOM found for ${assembly}`);
      continue;
    }

    const chosenRev = p.revision?.trim() || rows.map((x) => x.revision).sort().slice(-1)[0];
    const lines = rows.filter((x) => x.revision === chosenRev);
    if (lines.length === 0) {
      errors.push(`No BOM lines found for ${assembly} rev ${chosenRev}`);
      continue;
    }

    const requirements = new Map<string, number>();
    const shortages: Array<{ part: string; required: number; available: number; shortage: number }> = [];
    let maxBuildable = Number.MAX_SAFE_INTEGER;

    for (const l of lines) {
      const qpb = Number(l.qty_per_board || 0);
      if (qpb <= 0) continue;
      const req = qpb * qty;
      const avail = available.get(l.component_pn) || 0;
      requirements.set(l.component_pn, req);
      if (req > avail) shortages.push({ part: l.component_pn, required: req, available: avail, shortage: req - avail });
      maxBuildable = Math.min(maxBuildable, Math.floor(avail / qpb));
    }

    computed.push({
      run_id: String(p.run_id || `${assembly}__${chosenRev}__${computed.length}`),
      run_label: String(p.run_label || `${assembly} rev ${chosenRev}`),
      assembly_pn: assembly,
      revision: chosenRev,
      qty,
      priority,
      max_buildable: maxBuildable === Number.MAX_SAFE_INTEGER ? 0 : maxBuildable,
      can_complete: shortages.length === 0,
      shortages,
      requirements,
    });
  }

  const recommended = findBestSubset(computed, available);
  const recommendedKey = new Set(recommended.map((x) => x.run_id));

  // Build per-BOM shortage rows so each row shows which run it belongs to
  const perBomShortages: Array<{ part: string; shortage: number; required: number; available: number; bom_run: string; priority: number }> = [];
  for (const c of computed) {
    const bomLabel = c.run_label || `${c.assembly_pn} rev ${c.revision}`;
    for (const s of c.shortages) {
      perBomShortages.push({ part: s.part, shortage: s.shortage, required: s.required, available: s.available, bom_run: bomLabel, priority: c.priority });
    }
  }

  return {
    ok: true,
    objective: "complete_run_first",
    source: "supabase",
    plans: computed.map((c) => ({
      run_id: c.run_id,
      run_label: c.run_label,
      assembly_pn: c.assembly_pn,
      revision: c.revision,
      qty: c.qty,
      priority: c.priority,
      can_complete: c.can_complete,
      max_buildable: c.max_buildable,
      shortages: c.shortages,
      in_recommended: recommendedKey.has(c.run_id),
    })),
    recommended: recommended.map((c) => ({
      run_id: c.run_id,
      run_label: c.run_label,
      assembly_pn: c.assembly_pn,
      revision: c.revision,
      qty: c.qty,
      priority: c.priority,
    })),
    summary: {
      requested_runs: computed.length,
      requested_boards: computed.reduce((s, x) => s + x.qty, 0),
      individually_completable_runs: computed.filter((x) => x.can_complete).length,
      recommended_complete_runs: recommended.length,
      recommended_boards: recommended.reduce((s, x) => s + x.qty, 0),
    },
    top_shortages: perBomShortages
      .sort((a, b) => b.shortage - a.shortage),
    errors,
  };
}

function analyzeFromCsv(plans: PlanInput[]) {
  const root = process.cwd();
  const bomRows = parseCsv(path.join(root, "data", "bom.csv"));
  const invRows = parseCsv(path.join(root, "data", "inventory_lots.csv"));

  const available = new Map<string, number>();
  for (let i = 1; i < invRows.length; i++) {
    const part = invRows[i][1];
    const onHand = Number(invRows[i][2] || 0);
    const reserved = Number(invRows[i][3] || 0);
    available.set(part, (available.get(part) || 0) + Math.max(0, onHand - reserved));
  }

  const byAssemblyRev = new Map<string, Array<{ component_pn: string; qty_per_board: number }>>();
  for (let i = 1; i < bomRows.length; i++) {
    const assembly = bomRows[i][0];
    const revision = bomRows[i][1];
    const key = `${assembly}__${revision}`;
    if (!byAssemblyRev.has(key)) byAssemblyRev.set(key, []);
    byAssemblyRev.get(key)!.push({ component_pn: bomRows[i][2], qty_per_board: Number(bomRows[i][3] || 0) });
  }

  const computed: PlanComputed[] = [];
  for (const p of plans) {
    const assembly = String(p.assembly_pn || "").trim();
    const qty = Math.max(1, Math.floor(Number(p.qty || 0)));
    const priority = Number(p.priority || 1);
    const rev = String(p.revision || "A");
    const lines = byAssemblyRev.get(`${assembly}__${rev}`) || [];
    if (lines.length === 0) continue;

    const requirements = new Map<string, number>();
    const shortages: Array<{ part: string; required: number; available: number; shortage: number }> = [];
    let maxBuildable = Number.MAX_SAFE_INTEGER;

    for (const l of lines) {
      const qpb = Number(l.qty_per_board || 0);
      if (qpb <= 0) continue;
      const req = qpb * qty;
      const avail = available.get(l.component_pn) || 0;
      requirements.set(l.component_pn, req);
      if (req > avail) shortages.push({ part: l.component_pn, required: req, available: avail, shortage: req - avail });
      maxBuildable = Math.min(maxBuildable, Math.floor(avail / qpb));
    }

    computed.push({
      run_id: String(p.run_id || `${assembly}__${rev}__${computed.length}`),
      run_label: String(p.run_label || `${assembly} rev ${rev}`),
      assembly_pn: assembly,
      revision: rev,
      qty,
      priority,
      max_buildable: maxBuildable === Number.MAX_SAFE_INTEGER ? 0 : maxBuildable,
      can_complete: shortages.length === 0,
      shortages,
      requirements,
    });
  }

  const recommended = findBestSubset(computed, available);
  const recommendedKey = new Set(recommended.map((x) => x.run_id));

  const csvPerBomShortages: Array<{ part: string; shortage: number; required: number; available: number; bom_run: string; priority: number }> = [];
  for (const c of computed) {
    const bomLabel = c.run_label || `${c.assembly_pn} rev ${c.revision}`;
    for (const s of c.shortages) {
      csvPerBomShortages.push({ part: s.part, shortage: s.shortage, required: s.required, available: s.available, bom_run: bomLabel, priority: c.priority });
    }
  }

  return {
    ok: true,
    objective: "complete_run_first",
    source: "csv-fallback",
    plans: computed.map((c) => ({
      run_id: c.run_id,
      run_label: c.run_label,
      assembly_pn: c.assembly_pn,
      revision: c.revision,
      qty: c.qty,
      priority: c.priority,
      can_complete: c.can_complete,
      max_buildable: c.max_buildable,
      shortages: c.shortages,
      in_recommended: recommendedKey.has(c.run_id),
    })),
    recommended: recommended.map((c) => ({
      run_id: c.run_id,
      run_label: c.run_label,
      assembly_pn: c.assembly_pn,
      revision: c.revision,
      qty: c.qty,
      priority: c.priority,
    })),
    summary: {
      requested_runs: computed.length,
      requested_boards: computed.reduce((s, x) => s + x.qty, 0),
      individually_completable_runs: computed.filter((x) => x.can_complete).length,
      recommended_complete_runs: recommended.length,
      recommended_boards: recommended.reduce((s, x) => s + x.qty, 0),
    },
    top_shortages: csvPerBomShortages
      .sort((a, b) => b.shortage - a.shortage),
    errors: [],
  };
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const plans = Array.isArray(body?.plans) ? (body.plans as PlanInput[]) : [];
    if (plans.length === 0) {
      return NextResponse.json({ error: "plans required" }, { status: 400 });
    }

    try {
      const result = await analyzeFromSupabase(plans);
      return NextResponse.json(result);
    } catch {
      const result = analyzeFromCsv(plans);
      return NextResponse.json(result);
    }
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Unexpected error";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
