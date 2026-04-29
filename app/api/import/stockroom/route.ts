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

    // Skip header row if present
    const dataRows = rows.length > 0 && rows[0][0]?.toLowerCase().includes("item") ? rows.slice(1) : rows;

    // CSV columns: A=Item (part number), B=Bin Number, C=Inventory Number (lot), D=On Hand (qty), E=Available, F=Location, G=Expiration
    const lots = dataRows
      .map((r) => ({
        component_pn: (r[0] || "").trim(),
        bin_location: (r[1] || "").trim(),
        lot_number: (r[2] || "").trim(),
        qty: Math.floor(Number(r[3] || 0)),
        status: "ACTIVE",
      }))
      .filter((r) => r.component_pn && r.lot_number);

    // Full snapshot replace
    const { error: delErr } = await supabase.from("stockroom_lots").delete().neq("lot_number", "__never__");
    if (delErr) throw delErr;

    // Insert in batches of 500 to avoid payload limits
    for (let i = 0; i < lots.length; i += 500) {
      const batch = lots.slice(i, i + 500);
      const { error } = await supabase.from("stockroom_lots").insert(batch);
      if (error) throw new Error(`Insert batch ${Math.floor(i / 500) + 1} failed: ${error.message}`);
    }

    return NextResponse.json({ ok: true, imported: lots.length, mode: "replace_snapshot" });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Import failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
