/**
 * Adapter interface for any pricing/availability provider
 * (TrustedParts, Octopart/Nexar, SiliconExpert, custom CSV, etc).
 *
 * The optimizer speaks ONLY to this interface. Anything provider-specific
 * (auth, pagination, rate limits, field mapping) lives inside the adapter.
 */

import type { Offer } from "../types";

export type ProviderQuery = {
  part_number: string;
  qty: number;
  /** Optional hint: preferred currency, region, etc. */
  region?: string;
  currency?: string;
};

export type ProviderLookupResult = {
  part_number: string;
  offers: Offer[];
  /** Optional manufacturer / description metadata from the provider. */
  manufacturer?: string;
  description?: string;
  /** Raw provider identifier if we need to trace back. */
  provider_part_id?: string;
  /** Optional xref suggestions from provider data. */
  provider_alternates?: string[];
};

export type ProviderHealth = {
  ok: boolean;
  provider: string;
  mode: "mock" | "live";
  note?: string;
  latency_ms?: number;
};

export interface PricingAdapter {
  /** Short id, e.g. "mock", "trustedparts". */
  readonly name: string;
  /** "mock" when using fixtures, "live" when hitting the real API. */
  readonly mode: "mock" | "live";
  /** Bulk lookup. Providers SHOULD batch internally when possible. */
  lookup(queries: ProviderQuery[]): Promise<ProviderLookupResult[]>;
  /** Cheap health probe for diagnostics. */
  health(): Promise<ProviderHealth>;
}
