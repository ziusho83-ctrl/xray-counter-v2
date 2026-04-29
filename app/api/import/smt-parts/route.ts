import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";
import { isAdmin } from "@/lib/admin";

/**
 * SMT Parts list import/read.
 * POST: Upload CSV of SMT part numbers (replaces existing list)
 * GET:  Return all SMT part numbers
 */

export async function POST(req: NextRequest) {
  try {
    if (!isAdmin(req)) return NextResponse.json({ error: "Admin only" }, { status: 403 });
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const form = await req.formData();
    const file = form.get("file") as File | null;
    if (!file) return NextResponse.json({ error: "file required" }, { status: 400 });

    const text = await file.text();
    const lines = text.split(/\r?\n/).filter(Boolean);
    
    // Detect header
    const firstLine = lines[0].trim().toLowerCase();
    const hasHeader = firstLine.includes("part") || firstLine.includes("smt") || firstLine.includes("number");
    const dataLines = hasHeader ? lines.slice(1) : lines;

    const parts = [...new Set(
      dataLines
        .map((l) => l.split(",")[0].trim())
        .filter(Boolean)
    )];

    if (parts.length === 0) {
      return NextResponse.json({ error: "No parts found in CSV" }, { status: 400 });
    }

    // Clear existing
    const { error: delErr } = await supabase.from("smt_parts").delete().gte("component_pn", "");
    if (delErr) throw new Error(`Failed to clear smt_parts: ${delErr.message}`);

    // Insert in batches
    const BATCH = 500;
    for (let i = 0; i < parts.length; i += BATCH) {
      const batch = parts.slice(i, i + BATCH).map((pn) => ({ component_pn: pn }));
      const { error } = await supabase.from("smt_parts").upsert(batch, { onConflict: "component_pn" });
      if (error) throw new Error(`Failed to insert smt_parts batch: ${error.message}`);
    }

    return NextResponse.json({ ok: true, imported: parts.length });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "SMT parts import failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function GET() {
  try {
    if (!supabase) return NextResponse.json({ items: [] });

    // Paginate to get all SMT parts
    const PAGE = 1000;
    const allParts: Array<{ component_pn: string }> = [];
    let offset = 0;
    while (true) {
      const { data, error } = await supabase
        .from("smt_parts")
        .select("component_pn")
        .range(offset, offset + PAGE - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      allParts.push(...data);
      if (data.length < PAGE) break;
      offset += PAGE;
    }

    return NextResponse.json({ items: allParts, count: allParts.length });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Failed to fetch SMT parts";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
