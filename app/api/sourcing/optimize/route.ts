import { NextResponse } from "next/server";
import { optimize } from "@/lib/sourcing/optimize";
import type { BomLine, PricingTable, XrefTable } from "@/lib/sourcing/types";
import { getPricingAdapter } from "@/lib/sourcing/adapters";

export const runtime = "nodejs";

function coerceBom(input: unknown): BomLine[] {
  if (!Array.isArray(input)) return [];
  const out: BomLine[] = [];
  for (const row of input) {
    if (!row || typeof row !== "object") continue;
    const pn = String((row as any).part_number ?? (row as any).pn ?? "").trim();
    const qtyRaw = (row as any).qty ?? (row as any).quantity;
    const qty = Number(qtyRaw);
    if (!pn || !Number.isFinite(qty) || qty <= 0) continue;
    out.push({ part_number: pn, qty: Math.floor(qty) });
  }
  return out;
}

export async function POST(req: Request) {
  try {
    const body = await req.json().catch(() => ({}));
    const bom = coerceBom(body?.bom);
    if (bom.length === 0) {
      return NextResponse.json(
        { error: "Empty or invalid BOM. Expected { bom: [{ part_number, qty }] }." },
        { status: 400 },
      );
    }

    const adapter = getPricingAdapter();
    const queries = bom.map((b) => ({ part_number: b.part_number, qty: b.qty }));

    // Request offers for every BOM line in one bulk call. For providers that
    // can't batch, the adapter fans out internally.
    const lookups = await adapter.lookup(queries);

    // Collect xref suggestions from the adapter, merged with any requested
    // alternates. For the mock adapter this is just the bundled xref JSON.
    const pricing: PricingTable = {};
    const xref: XrefTable = {};
    for (const r of lookups) {
      pricing[r.part_number] = r.offers || [];
      if (r.provider_alternates && r.provider_alternates.length > 0) {
        xref[r.part_number] = r.provider_alternates;
      }
    }

    // Also resolve offers for any alternate parts we might need to fall back to.
    // (The mock adapter already has them cached; for a live adapter this is a
    // second round of lookups kept small by only targeting unsourced lines.)
    const unsourced = bom.filter(
      (b) => !pricing[b.part_number] || pricing[b.part_number].length === 0,
    );
    if (unsourced.length > 0) {
      const altQueries = new Map<string, number>();
      for (const u of unsourced) {
        const alts = xref[u.part_number] || [];
        for (const a of alts) {
          if (!(a in pricing)) altQueries.set(a, u.qty);
        }
      }
      if (altQueries.size > 0) {
        const altLookups = await adapter.lookup(
          Array.from(altQueries.entries()).map(([pn, qty]) => ({ part_number: pn, qty })),
        );
        for (const r of altLookups) {
          pricing[r.part_number] = r.offers || [];
        }
      }
    }

    const result = optimize(bom, pricing, xref);
    return NextResponse.json({
      ...result,
      provider: {
        name: adapter.name,
        mode: adapter.mode,
      },
    });
  } catch (e: any) {
    return NextResponse.json(
      { error: e?.message || "Unexpected error" },
      { status: 500 },
    );
  }
}

export async function GET() {
  const adapter = getPricingAdapter();
  return NextResponse.json({
    ok: true,
    version: "0.3",
    provider: { name: adapter.name, mode: adapter.mode },
    note: "POST { bom: [{ part_number, qty }] } to get a sourcing plan.",
  });
}
