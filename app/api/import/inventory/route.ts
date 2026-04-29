import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";
import { isAdmin } from "@/lib/admin";

function parseCsv(raw: string): string[][] {
  return raw
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => line.split(",").map((x) => x.trim()));
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

    // expected: part number, lot number, qty, date scan
    const dataRows = rows;

    const lots = dataRows
      .map((r) => ({
        component_pn: r[0],
        lot_id: r[1],
        qty_on_hand: Number(r[2] || 0),
        qty_reserved: 0,
        location: r[4]?.trim() || "LINE",
        status: "ACTIVE",
      }))
      .filter((r) => r.component_pn && r.lot_id);

    // Treat upload as full snapshot: replace inventory table content
    const { error: delErr } = await supabase.from("inventory_lots").delete().neq("lot_id", "__never__");
    if (delErr) throw delErr;

    const { error } = await supabase.from("inventory_lots").insert(lots);
    if (error) throw error;

    return NextResponse.json({ ok: true, imported: lots.length, mode: "replace_snapshot" });
  } catch (e: any) {
    return NextResponse.json({ error: e.message || "Import failed" }, { status: 500 });
  }
}
