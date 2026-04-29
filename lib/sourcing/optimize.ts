import type {
  Offer,
  PricingTable,
  XrefTable,
  BomLine,
  PlanRow,
  OptimizeResult,
  LineStatus,
} from "./types";
import { findAlternates } from "./xref";

/**
 * Rank offers for a given quantity.
 * Qualified = offers where stock >= max(qty, moq).
 * Partial   = offers with some stock but not enough.
 * Ordered by extended price asc, then lead time asc.
 */
function rankOffers(
  qty: number,
  offers: Offer[],
): { ranked: Offer[]; status: LineStatus } {
  const qualified: Offer[] = [];
  const partial: Offer[] = [];
  for (const o of offers) {
    const effQty = Math.max(qty, o.moq || 1);
    if (o.stock >= effQty) qualified.push(o);
    else if (o.stock > 0) partial.push(o);
  }

  const rankKey = (o: Offer): [number, number] => {
    const effQty = Math.max(qty, o.moq || 1);
    return [o.unit_price * effQty, o.lead_time_days];
  };
  const cmp = (a: Offer, b: Offer) => {
    const [pa, la] = rankKey(a);
    const [pb, lb] = rankKey(b);
    if (pa !== pb) return pa - pb;
    return la - lb;
  };

  qualified.sort(cmp);
  partial.sort(cmp);

  if (qualified.length > 0) {
    return { ranked: [...qualified, ...partial], status: "ok" };
  }
  if (partial.length > 0) {
    return { ranked: partial, status: "partial_stock" };
  }
  return { ranked: [], status: "no_source" };
}

function tryPart(
  pn: string,
  qty: number,
  pricing: PricingTable,
): { ranked: Offer[]; status: LineStatus } {
  const offers = pricing[pn];
  if (!offers || offers.length === 0) {
    return { ranked: [], status: "no_source" };
  }
  return rankOffers(qty, offers);
}

export function optimize(
  bom: BomLine[],
  pricing: PricingTable,
  xref: XrefTable = {},
): OptimizeResult {
  const rows: PlanRow[] = [];

  for (const { part_number: requestedPn, qty } of bom) {
    const tried: string[] = [];
    tried.push(requestedPn);

    let { ranked, status } = tryPart(requestedPn, qty, pricing);
    let pickedPn = requestedPn;
    let subReason = "";
    let note = "";

    if (status !== "ok") {
      const alternates = findAlternates(requestedPn, xref);
      let bestCandidate:
        | { pn: string; ranked: Offer[]; status: LineStatus }
        | null = null;
      for (const altPn of alternates) {
        tried.push(altPn);
        const { ranked: altRanked, status: altStatus } = tryPart(altPn, qty, pricing);
        if (altStatus === "ok") {
          bestCandidate = { pn: altPn, ranked: altRanked, status: "ok" };
          break;
        }
        if (altStatus === "partial_stock" && bestCandidate === null) {
          bestCandidate = { pn: altPn, ranked: altRanked, status: "partial_stock" };
        }
      }

      if (bestCandidate) {
        if (status === "no_source") subReason = "original part not in pricing table";
        else if (status === "partial_stock") subReason = "original part had insufficient stock";
        pickedPn = bestCandidate.pn;
        ranked = bestCandidate.ranked;
        status = bestCandidate.status;
      }
    }

    if (ranked.length === 0) {
      const offersForPn = pricing[requestedPn];
      const hasPricing = !!(offersForPn && offersForPn.length > 0);
      const hasXref = findAlternates(requestedPn, xref).length > 0;
      if (!hasPricing && !hasXref) {
        note = "part not in pricing table; no xref alternates";
      } else if (!hasPricing) {
        note = "part not in pricing table; xref alternates also unsourceable";
      } else {
        note = "no sourceable offer for requested part or alternates";
      }
      rows.push({
        requested_pn: requestedPn,
        picked_pn: requestedPn,
        qty,
        status: "no_source",
        substituted_from: "",
        sub_reason: "",
        best: null,
        alternates: [],
        note,
        tried_parts: tried,
        ext_price: 0,
      });
      continue;
    }

    const best = ranked[0];
    const alternates = ranked.slice(1);
    if (status === "partial_stock") {
      note = "no source has full stock; best is partial";
    }
    const substitutedFrom = pickedPn !== requestedPn ? requestedPn : "";
    const effQty = Math.max(qty, best.moq || 1);
    const extPrice = Math.round(best.unit_price * effQty * 10000) / 10000;

    rows.push({
      requested_pn: requestedPn,
      picked_pn: pickedPn,
      qty,
      status,
      substituted_from: substitutedFrom,
      sub_reason: subReason,
      best,
      alternates,
      note,
      tried_parts: tried,
      ext_price: extPrice,
    });
  }

  // Summary
  let fullySourced = 0;
  let partialStock = 0;
  let noSource = 0;
  let substituted = 0;
  let totalCost = 0;
  let criticalLead = 0;
  for (const r of rows) {
    if (r.status === "ok") fullySourced++;
    else if (r.status === "partial_stock") partialStock++;
    else noSource++;
    if (r.substituted_from) substituted++;
    if (r.best) {
      const effQty = Math.max(r.qty, r.best.moq || 1);
      totalCost += r.best.unit_price * effQty;
      if (r.best.lead_time_days > criticalLead) criticalLead = r.best.lead_time_days;
    }
  }

  return {
    rows,
    summary: {
      lines_total: rows.length,
      fully_sourced: fullySourced,
      partial_stock: partialStock,
      no_source: noSource,
      substituted,
      total_cost: Math.round(totalCost * 100) / 100,
      critical_lead_days: criticalLead,
    },
  };
}
