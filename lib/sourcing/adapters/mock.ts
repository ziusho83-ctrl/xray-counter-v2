import type {
  PricingAdapter,
  ProviderHealth,
  ProviderLookupResult,
  ProviderQuery,
} from "./types";
import type { PricingTable, XrefTable } from "../types";
import pricingRaw from "../mock_pricing.json";
import xrefRaw from "../mock_xref.json";
import { loadXrefFromRaw } from "../xref";

/**
 * Mock adapter — serves pricing + xref data from bundled JSON fixtures.
 * This is what the /api/sourcing/optimize endpoint uses by default and
 * what runs in dev / preview when no live credentials are configured.
 */
export class MockPricingAdapter implements PricingAdapter {
  readonly name = "mock";
  readonly mode = "mock" as const;

  private pricing: PricingTable;
  private xref: XrefTable;

  constructor() {
    this.pricing = pricingRaw as unknown as PricingTable;
    this.xref = loadXrefFromRaw(xrefRaw as unknown as Record<string, unknown>);
  }

  async lookup(queries: ProviderQuery[]): Promise<ProviderLookupResult[]> {
    const out: ProviderLookupResult[] = [];
    for (const q of queries) {
      const offers = this.pricing[q.part_number] || [];
      const alternates = this.xref[q.part_number] || [];
      out.push({
        part_number: q.part_number,
        offers: [...offers],
        provider_alternates: alternates.length > 0 ? [...alternates] : undefined,
      });
    }
    return out;
  }

  async health(): Promise<ProviderHealth> {
    return {
      ok: true,
      provider: this.name,
      mode: this.mode,
      note: "mock fixtures loaded from lib/sourcing/mock_pricing.json",
      latency_ms: 0,
    };
  }
}
