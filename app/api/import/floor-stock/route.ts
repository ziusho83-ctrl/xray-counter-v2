import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";
import { isAdmin } from "@/lib/admin";

function parseCsv(raw: string): string[][] {
  return raw
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => line.split(",").map((x) => x.trim()));
}

const PAGE_SIZE = 1000;

export async function GET() {
  try {
    if (!supabase) return NextResponse.json({ items: [], error: "Supabase not configured" });

    // Paginate to avoid Supabase 1000-row default limit
    const allItems: Array<{ component_pn: string; location: string | null }> = [];
    let offset = 0;
    while (true) {
      const { data, error } = await supabase
        .from("floor_stock_parts")
        .select("component_pn, location")
        .order("component_pn")
        .range(offset, offset + PAGE_SIZE - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      allItems.push(...data);
      if (data.length < PAGE_SIZE) break;
      offset += PAGE_SIZE;
    }

    return NextResponse.json({ items: allItems });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Failed to load floor stock";
    return NextResponse.json({ items: [], error: message });
  }
}

export async function POST(req: NextRequest) {
  try {
    if (!isAdmin(req)) return NextResponse.json({ error: "Admin only" }, { status: 403 });
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const form = await req.formData();
    const file = form.get("file") as File | null;
    if (!file) return NextResponse.json({ error: "file required" }, { status: 400 });

    const text = await file.text();
    const rows = parseCsv(text);

    // Detect header row
    const hasHeader = rows.length > 0 && rows[0][0]?.toLowerCase().includes("part") || rows[0][0]?.toLowerCase().includes("component") || rows[0][0]?.toLowerCase().includes("item");
    const dataRows = hasHeader ? rows.slice(1) : rows;

    // CSV columns: A=Part#, B=Location
    const parts = dataRows
      .map((r) => ({
        component_pn: (r[0] || "").trim(),
        location: (r[1] || "").trim() || null,
      }))
      .filter((r) => r.component_pn);

    // Deduplicate by component_pn
    const seen = new Set<string>();
    const unique = parts.filter((p) => {
      if (seen.has(p.component_pn)) return false;
      seen.add(p.component_pn);
      return true;
    });

    // Full snapshot replace
    const { error: delErr } = await supabase.from("floor_stock_parts").delete().neq("component_pn", "__never__");
    if (delErr) throw delErr;

    // Insert in batches
    for (let i = 0; i < unique.length; i += 500) {
      const batch = unique.slice(i, i + 500);
      const { error } = await supabase.from("floor_stock_parts").insert(batch);
      if (error) throw new Error(`Insert batch ${Math.floor(i / 500) + 1} failed: ${error.message}`);
    }

    return NextResponse.json({ ok: true, imported: unique.length, mode: "replace_snapshot" });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Import failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function DELETE(req: NextRequest) {
  try {
    if (!isAdmin(req)) return NextResponse.json({ error: "Admin only" }, { status: 403 });
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const body = await req.json();
    const pn = body?.component_pn;
    if (!pn) return NextResponse.json({ error: "component_pn required" }, { status: 400 });

    const { error, count } = await supabase.from("floor_stock_parts").delete().eq("component_pn", pn);
    if (error) throw error;

    return NextResponse.json({ ok: true, deleted: count || 1 });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Delete failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
