import { NextRequest, NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";
import { isAdmin } from "@/lib/admin";

export async function POST(req: NextRequest) {
  try {
    if (!isAdmin(req)) return NextResponse.json({ error: "Admin only" }, { status: 403 });
    if (!supabase) return NextResponse.json({ error: "Supabase not configured" }, { status: 500 });

    const { assembly_pn, revision } = await req.json();
    if (!assembly_pn || !revision) {
      return NextResponse.json({ error: "assembly_pn and revision required" }, { status: 400 });
    }

    const { count, error: countErr } = await supabase
      .from("bom_lines")
      .select("component_pn", { count: "exact", head: true })
      .eq("assembly_pn", assembly_pn)
      .eq("revision", revision);
    if (countErr) throw countErr;

    const { error: delErr } = await supabase
      .from("bom_lines")
      .delete()
      .eq("assembly_pn", assembly_pn)
      .eq("revision", revision);
    if (delErr) throw delErr;

    await supabase
      .from("assemblies")
      .update({ active: false })
      .eq("assembly_pn", assembly_pn)
      .eq("revision", revision);

    return NextResponse.json({ ok: true, deleted: count || 0 });
  } catch (e: any) {
    return NextResponse.json({ error: e.message || "Delete failed" }, { status: 500 });
  }
}
