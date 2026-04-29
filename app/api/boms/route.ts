import { NextResponse } from "next/server";
import { supabase } from "@/lib/supabase";

export async function GET() {
  try {
    if (!supabase) return NextResponse.json({ items: [] });

    // Paginate to get all bom_lines (Supabase default limit is 1000)
    const allRows: Array<{ assembly_pn: string; revision: string; component_pn: string }> = [];
    const pageSize = 1000;
    let from = 0;
    while (true) {
      const { data, error } = await supabase
        .from("bom_lines")
        .select("assembly_pn, revision, component_pn")
        .range(from, from + pageSize - 1);
      if (error) throw error;
      if (!data || data.length === 0) break;
      allRows.push(...data);
      if (data.length < pageSize) break;
      from += pageSize;
    }

    const map = new Map<string, { assembly_pn: string; revision: string; line_count: number; bom_type: string }>();
    for (const r of allRows) {
      const key = `${r.assembly_pn}__${r.revision}`;
      if (!map.has(key)) map.set(key, { assembly_pn: r.assembly_pn, revision: r.revision, line_count: 0, bom_type: "PWB" });
      map.get(key)!.line_count += 1;
    }

    // Fetch bom_type from assemblies table
    const asmTypes = new Map<string, string>();
    let asmFrom = 0;
    while (true) {
      const { data, error } = await supabase
        .from("assemblies")
        .select("assembly_pn, revision, bom_type")
        .range(asmFrom, asmFrom + pageSize - 1);
      if (error) break;
      if (!data || data.length === 0) break;
      for (const a of data) asmTypes.set(`${a.assembly_pn}__${a.revision}`, a.bom_type || "PWB");
      if (data.length < pageSize) break;
      asmFrom += pageSize;
    }
    for (const [key, item] of map) {
      item.bom_type = asmTypes.get(key) || "PWB";
    }

    const items = Array.from(map.values()).sort((a, b) => a.assembly_pn.localeCompare(b.assembly_pn));
    return NextResponse.json({ items });
  } catch {
    return NextResponse.json({ items: [] });
  }
}
