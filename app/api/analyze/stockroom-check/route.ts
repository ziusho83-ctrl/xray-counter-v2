import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";

type ShortageInput = { part: string; shortage: number; bom_run: string };

// Supabase/PostgREST defaults to 1000 rows per query.
// For large MPS runs we must paginate to get all stockroom lots.
const PAGE_SIZE = 1000;

async function fetchAllStockroomLots(
  client: NonNullable<typeof supabase>,
  parts: string[]
) {
  // Batch the .in() filter to avoid URL-length limits (~300 parts per batch)
  const IN_BATCH = 300;
  const allLots: Array<{ component_pn: string; lot_number: string; qty: number; bin_location: string }> = [];

  for (let i = 0; i < parts.length; i += IN_BATCH) {
    const batch = parts.slice(i, i + IN_BATCH);
    let offset = 0;
    let hasMore = true;

    while (hasMore) {
      const { data, error } = await client
        .from("stockroom_lots")
        .select("component_pn, lot_number, qty, bin_location")
        .in("component_pn", batch)
        .eq("status", "ACTIVE")
        .gt("qty", 0)
        .range(offset, offset + PAGE_SIZE - 1);
      if (error) throw error;
      const rows = data || [];
      allLots.push(...rows);
      hasMore = rows.length === PAGE_SIZE;
      offset += PAGE_SIZE;
    }
  }

  return allLots;
}

export async function POST(req: NextRequest) {
  try {
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const body = await req.json();
    const shortages: ShortageInput[] = body?.shortages || [];
    if (shortages.length === 0) return NextResponse.json({ results: [] });

    const parts = [...new Set(shortages.map((s) => s.part))];

    // Fetch ALL stockroom lots, paginating past the 1000-row default limit
    // and batching the .in() filter to stay within URL-length limits.
    const filteredStockroom = await fetchAllStockroomLots(supabase, parts);

    // Group by component_pn
    const stockroomByPart = new Map<string, Array<{ bin_location: string; lot_number: string; qty: number }>>();
    for (const sl of filteredStockroom) {
      if (!stockroomByPart.has(sl.component_pn)) stockroomByPart.set(sl.component_pn, []);
      stockroomByPart.get(sl.component_pn)!.push({
        bin_location: sl.bin_location,
        lot_number: sl.lot_number,
        qty: sl.qty,
      });
    }

    const INSPECTION_PATTERN = /inspection/i;

    const results = shortages
      .map((s) => {
        const locations = stockroomByPart.get(s.part) || [];
        let confirmed_qty = 0;
        let inspection_qty = 0;
        for (const l of locations) {
          if (INSPECTION_PATTERN.test(l.bin_location)) {
            inspection_qty += l.qty;
          } else {
            confirmed_qty += l.qty;
          }
        }
        return {
          part: s.part,
          shortage_qty: s.shortage,
          stockroom_available: confirmed_qty + inspection_qty, // total (kept for backward compat)
          confirmed_qty,
          inspection_qty,
          stockroom_locations: locations,
          bom_run: s.bom_run,
        };
      })
      .sort((a, b) => b.shortage_qty - a.shortage_qty);

    return NextResponse.json({ results });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Stockroom check failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
