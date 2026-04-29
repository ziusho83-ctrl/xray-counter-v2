import { NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";

export async function GET() {
  try {
    if (!supabase) return NextResponse.json({ items: [] });

    const { data, error } = await supabase
      .from("build_checks")
      .select("check_id, assembly_pn, revision, build_qty, can_run, shortage_count, max_buildable, created_at")
      .order("created_at", { ascending: false })
      .limit(20);

    if (error) throw error;
    return NextResponse.json({ items: data || [] });
  } catch {
    return NextResponse.json({ items: [] });
  }
}
