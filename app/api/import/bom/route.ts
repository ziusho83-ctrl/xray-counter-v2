import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";
import { isAdmin } from "@/lib/admin";

function parseCsv(raw: string): string[][] {
  return raw
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => line.split(",").map((x) => x.trim()));
}

/** Compare two revisions. Returns positive if a > b, negative if a < b, 0 if equal.
 *  Auto-detects numeric vs alphabetical revisions. */
function compareRevisions(a: string, b: string): number {
  const numA = Number(a);
  const numB = Number(b);
  // If both parse as numbers, compare numerically
  if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
  // Otherwise compare alphabetically (case-insensitive)
  return a.toLowerCase().localeCompare(b.toLowerCase());
}

export async function POST(req: NextRequest) {
  try {
    if (!isAdmin(req)) return NextResponse.json({ error: "Admin only" }, { status: 403 });
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const form = await req.formData();
    const assembly = String(form.get("assembly_pn") || "").trim();
    const revision = String(form.get("revision") || "").trim();
    const file = form.get("file") as File | null;
    const dryRun = String(form.get("dry_run") || "") === "1";

    if (!assembly || !revision || !file) {
      return NextResponse.json({ error: "assembly_pn, revision, file required" }, { status: 400 });
    }

    const text = await file.text();
    const rows = parseCsv(text);
    const dataRows = rows[0]?.[0]?.toLowerCase().includes("component") ? rows.slice(1) : rows;

    const parsed = dataRows
      .map((r) => ({ component_pn: String(r[0] || "").trim(), qty_per_board: Number(r[1] || 0) }))
      .filter((r) => r.component_pn && r.qty_per_board > 0);

    // Consolidate duplicate component rows in BOM CSV to avoid PK collisions.
    const qtyByPart = new Map<string, number>();
    for (const row of parsed) {
      qtyByPart.set(row.component_pn, (qtyByPart.get(row.component_pn) || 0) + row.qty_per_board);
    }

    const bom = Array.from(qtyByPart.entries()).map(([component_pn, qty_per_board]) => ({
      assembly_pn: assembly,
      revision,
      component_pn,
      qty_per_board,
    }));
    const duplicateRowsMerged = parsed.length - bom.length;

    // --- Revision control: block older revs, auto-remove superseded revs ---
    const { data: existingRevs, error: revErr } = await supabase
      .from("bom_lines")
      .select("revision")
      .eq("assembly_pn", assembly);
    if (revErr) throw revErr;

    const uniqueRevs = [...new Set((existingRevs || []).map((r) => String(r.revision).trim()))];
    const higherRevs = uniqueRevs.filter((r) => compareRevisions(r, revision) > 0);
    const lowerOrEqualRevs = uniqueRevs.filter((r) => compareRevisions(r, revision) < 0);
    const sameRev = uniqueRevs.filter((r) => compareRevisions(r, revision) === 0);

    if (higherRevs.length > 0) {
      return NextResponse.json({
        error: `Cannot import rev ${revision} — newer revision(s) already exist: ${higherRevs.join(", ")}. Only the latest revision is allowed.`,
      }, { status: 400 });
    }

    const { count: existingCount, error: cntErr } = await supabase
      .from("bom_lines")
      .select("component_pn", { count: "exact", head: true })
      .eq("assembly_pn", assembly)
      .eq("revision", revision);
    if (cntErr) throw cntErr;

    if (dryRun) {
      return NextResponse.json({
        ok: true,
        dryRun: true,
        existingCount: existingCount || 0,
        incomingCount: bom.length,
        duplicateRowsMerged,
        revisionControl: {
          incomingRev: revision,
          oldRevsToRemove: lowerOrEqualRevs.length > 0 ? lowerOrEqualRevs : undefined,
          replacingSameRev: sameRev.length > 0,
        },
      });
    }

    // Remove older revision BOM lines and assembly records
    for (const oldRev of lowerOrEqualRevs) {
      await supabase.from("bom_lines").delete().eq("assembly_pn", assembly).eq("revision", oldRev);
      await supabase.from("assemblies").delete().eq("assembly_pn", assembly).eq("revision", oldRev);
    }

    const { error: asmErr } = await supabase.from("assemblies").upsert({ assembly_pn: assembly, revision, active: true });
    if (asmErr) throw asmErr;

    const { error: delErr } = await supabase.from("bom_lines").delete().eq("assembly_pn", assembly).eq("revision", revision);
    if (delErr) throw delErr;

    if (bom.length > 0) {
      // Batch insert to avoid Supabase payload/row limits
      const batchSize = 500;
      for (let i = 0; i < bom.length; i += batchSize) {
        const batch = bom.slice(i, i + batchSize);
        const { error: insErr } = await supabase.from("bom_lines").insert(batch);
        if (insErr) throw new Error(`BOM insert batch ${Math.floor(i / batchSize) + 1} failed: ${insErr.message}`);
      }
    }

    // Verify insert actually persisted
    const { count: verifyCount } = await supabase
      .from("bom_lines")
      .select("component_pn", { count: "exact", head: true })
      .eq("assembly_pn", assembly)
      .eq("revision", revision);

    return NextResponse.json({
      ok: true,
      imported: bom.length,
      verified: verifyCount || 0,
      duplicateRowsMerged,
      replaced: (existingCount || 0) > 0,
      replacedCount: existingCount || 0,
      oldRevsRemoved: lowerOrEqualRevs.length > 0 ? lowerOrEqualRevs : undefined,
    });
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Import failed";
    if (message.includes("duplicate key value")) {
      return NextResponse.json(
        { error: "Duplicate component rows found in BOM file. Please merge duplicate component_pn lines or retry (auto-merge now supported)." },
        { status: 400 }
      );
    }
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
