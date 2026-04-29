import type { XrefTable } from "./types";

/**
 * Return ordered alternates for a given part number, excluding the part itself
 * and duplicates. Keys beginning with "_" in the source JSON are treated as
 * comments and ignored by the loader.
 */
export function findAlternates(pn: string, xref: XrefTable): string[] {
  if (!pn) return [];
  const direct = xref[pn] || [];
  const seen = new Set<string>([pn]);
  const ordered: string[] = [];
  for (const alt of direct) {
    const trimmed = (alt || "").trim();
    if (!trimmed || seen.has(trimmed)) continue;
    seen.add(trimmed);
    ordered.push(trimmed);
  }
  return ordered;
}

export function loadXrefFromRaw(raw: Record<string, unknown>): XrefTable {
  const out: XrefTable = {};
  for (const [k, v] of Object.entries(raw)) {
    if (k.startsWith("_")) continue;
    if (!Array.isArray(v)) continue;
    out[k.trim()] = v.map((x) => String(x).trim()).filter(Boolean);
  }
  return out;
}
