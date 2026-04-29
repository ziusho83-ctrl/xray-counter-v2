import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";
import { isAdmin } from "@/lib/admin";

/**
 * Master BOM Import — accepts pre-parsed BOM data as JSON batches.
 * Client-side parses the large CSV, then sends structured data here.
 *
 * POST body:
 *   { action: "clear" }                          — delete all bom_lines + assemblies
 *   { action: "batch", assemblies: [...], lines: [...] } — insert a batch
 *   { action: "finalize" }                       — return final counts
 *
 * GET: Return summary of what's currently in the BOM database
 */

export async function POST(req: NextRequest) {
  try {
    if (!isAdmin(req)) return NextResponse.json({ error: "Admin only" }, { status: 403 });
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const body = await req.json();
    const action = body?.action;

    if (action === "clear") {
      // Supabase requires at least one filter on delete.
      // .gte("assembly_pn", "") matches all rows.
      const { error: delLinesErr } = await supabase.from("bom_lines").delete().gte("assembly_pn", "");
      if (delLinesErr) throw new Error(`Failed to clear bom_lines: ${delLinesErr.message}`);

      const { error: delAsmErr } = await supabase.from("assemblies").delete().gte("assembly_pn", "");
      if (delAsmErr) throw new Error(`Failed to clear assemblies: ${delAsmErr.message}`);

      return NextResponse.json({ ok: true, action: "clear" });
    }

    if (action === "batch") {
      const assemblies: Array<{ assembly_pn: string; revision: string; active: boolean; bom_type?: string }> = body.assemblies || [];
      const lines: Array<{ assembly_pn: string; revision: string; component_pn: string; qty_per_board: number }> = body.lines || [];

      // Upsert assemblies
      if (assemblies.length > 0) {
        const { error } = await supabase.from("assemblies").upsert(assemblies, { onConflict: "assembly_pn,revision" });
        if (error) throw new Error(`Failed to upsert assemblies: ${error.message}`);
      }

      // Insert bom_lines
      if (lines.length > 0) {
        const { error } = await supabase.from("bom_lines").insert(lines);
        if (error) throw new Error(`Failed to insert bom_lines: ${error.message}`);
      }

      return NextResponse.json({ ok: true, action: "batch", assemblies: assemblies.length, lines: lines.length });
    }

    if (action === "finalize") {
      const { count: asmCount } = await supabase.from("assemblies").select("assembly_pn", { count: "exact", head: true });
      const { count: lineCount } = await supabase.from("bom_lines").select("component_pn", { count: "exact", head: true });
      const { count: pwbCount } = await supabase.from("assemblies").select("assembly_pn", { count: "exact", head: true }).eq("bom_type", "PWB");
      const { count: harnessCount } = await supabase.from("assemblies").select("assembly_pn", { count: "exact", head: true }).eq("bom_type", "HARNESS");
      return NextResponse.json({ ok: true, action: "finalize", assemblies: asmCount || 0, bom_lines: lineCount || 0, pwb_assemblies: pwbCount || 0, harness_assemblies: harnessCount || 0 });
    }

    return NextResponse.json({ error: "Unknown action. Expected: clear, batch, or finalize" }, { status: 400 });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Master BOM import failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function GET(req: NextRequest) {
  try {
    if (!isAdmin(req)) return NextResponse.json({ error: "Admin only" }, { status: 403 });
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const { count: asmCount, error: asmErr } = await supabase
      .from("assemblies")
      .select("assembly_pn", { count: "exact", head: true });
    if (asmErr) throw asmErr;

    const { count: lineCount, error: lineErr } = await supabase
      .from("bom_lines")
      .select("component_pn", { count: "exact", head: true });
    if (lineErr) throw lineErr;

    return NextResponse.json({
      assemblies: asmCount || 0,
      bom_lines: lineCount || 0,
    });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Failed to get BOM summary";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
