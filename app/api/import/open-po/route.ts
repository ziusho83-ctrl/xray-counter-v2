import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";
import { isAdmin } from "@/lib/admin";

function parseCsv(raw: string): string[][] {
  // Handle quoted fields with commas inside
  const rows: string[][] = [];
  const lines = raw.split(/\r?\n/).filter(Boolean);
  for (const line of lines) {
    const fields: string[] = [];
    let current = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQuotes = !inQuotes;
      } else if (ch === "," && !inQuotes) {
        fields.push(current.trim());
        current = "";
      } else {
        current += ch;
      }
    }
    fields.push(current.trim());
    rows.push(fields);
  }
  return rows;
}

function asNum(v: string | undefined): number {
  if (!v) return 0;
  const n = Number(v.replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function asStr(v: string | undefined): string {
  return (v ?? "").trim();
}

function asDate(v: string | undefined): string {
  if (!v) return "";
  const s = v.trim();
  // Try Excel serial date
  const n = Number(s);
  if (Number.isFinite(n) && n > 30000 && n < 100000) {
    const d = new Date(Math.round((n - 25569) * 86400000));
    if (!isNaN(d.getTime())) {
      return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
    }
  }
  return s;
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
    const dataRows =
      rows.length > 0 && rows[0][0]?.toLowerCase().includes("vendor")
        ? rows.slice(1)
        : rows;

    // CSV columns: 0=vendor, 1=po, 2=line, 3=item, 4=description, 5=status(skip), 6=qty_ordered, 7=qty_received, 8=qty_due, 9=promise_date, 10=original_promise
    const parsed = dataRows
      .map((r) => ({
        vendor: asStr(r[0]),
        po: asStr(r[1]),
        line: asStr(r[2]),
        item: asStr(r[3]),
        description: asStr(r[4]),
        qty_ordered: asNum(r[6]),
        qty_received: asNum(r[7]),
        qty_due: asNum(r[8]),
        promise_date: asDate(r[9]),
        original_promise: asDate(r[10]),
      }))
      .filter((r) => {
        if (!r.item) return false;
        if (r.item.startsWith("58420") || r.item.startsWith("61")) return false;
        if (r.qty_due <= 0) return false;
        return true;
      });

    // Full snapshot replace
    const { error: delErr } = await supabase
      .from("open_purchase_orders")
      .delete()
      .neq("id", -1);
    if (delErr) throw delErr;

    // Insert in batches of 500
    for (let i = 0; i < parsed.length; i += 500) {
      const batch = parsed.slice(i, i + 500);
      const { error } = await supabase.from("open_purchase_orders").insert(batch);
      if (error)
        throw new Error(
          `Insert batch ${Math.floor(i / 500) + 1} failed: ${error.message}`
        );
    }

    return NextResponse.json({
      ok: true,
      imported: parsed.length,
      mode: "replace_snapshot",
    });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Import failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

export async function GET() {
  try {
    if (!supabase)
      return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    // Paginate to avoid Supabase 1000-row default limit
    const allItems: Array<Record<string, unknown>> = [];
    const pageSize = 1000;
    let offset = 0;
    while (true) {
      const { data, error } = await supabase
        .from("open_purchase_orders")
        .select("*")
        .order("item", { ascending: true })
        .range(offset, offset + pageSize - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      allItems.push(...data);
      if (data.length < pageSize) break;
      offset += pageSize;
    }

    return NextResponse.json({ items: allItems });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Failed to fetch open POs";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
