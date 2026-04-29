import type { Offer } from "../types";
import type {
  PricingAdapter,
  ProviderHealth,
  ProviderLookupResult,
  ProviderQuery,
} from "./types";

type TrustedPartsSearchQuery = {
  SearchToken: string;
  Manufacturers?: string[];
};

type TrustedPartsApiRequest = {
  Queries: TrustedPartsSearchQuery[];
  CountryCode: string;
  CurrencyCode: string;
  InStockOnly: boolean;
  ExactMatch: boolean;
  UseCachedData: boolean;
  Distributors?: string[];
  LanguageCode: string;
  IsCrawler: boolean;
  UserAgent: string;
  SourceIp: string;
};

type TrustedPartsPrice = {
  Quantity: number;
  Amount: number;
  FormattedAmount?: string;
  Text?: string;
};

type TrustedPartsProductPricing = {
  CurrencyCode?: string;
  MinimumQuantity?: number | null;
  QuantityMultiple?: number | null;
  Prices?: TrustedPartsPrice[] | null;
};

type TrustedPartsProductPackageType = {
  PackageType?: string;
  MinimumOrderQuantity?: number | null;
};

type TrustedPartsStockInfo = {
  QuantityOnHand?: number | null;
  Availability?: string | null;
};

type TrustedPartsComplianceEntry = {
  Region?: string;
  IsCompliant?: boolean;
  Description?: string;
};

type TrustedPartsProductCompliance = {
  RoHS?: TrustedPartsComplianceEntry[] | null;
};

type TrustedPartsSearchApiLink = {
  Type?: string;
  Url?: string;
};

type TrustedPartsDistributorResult = {
  Description?: string;
  DistributorPartNumber?: string;
  Compliance?: TrustedPartsProductCompliance;
  Stock?: TrustedPartsStockInfo;
  Links?: TrustedPartsSearchApiLink[] | null;
  Pricing?: TrustedPartsProductPricing;
  Packaging?: TrustedPartsProductPackageType[] | null;
};

type TrustedPartsDistributor = {
  Id?: number;
  Name?: string;
  DistributorResults?: TrustedPartsDistributorResult[] | null;
};

type TrustedPartsPartSpecification = {
  Key?: string;
  Value?: string;
};

type TrustedPartsPartResult = {
  PartNumber?: string;
  Manufacturer?: string;
  ManufacturerId?: number;
  ProductUrl?: string;
  IsAffectedByTariff?: boolean;
  LifecycleRisk?: string | null;
  SupplyChainRisk?: string | null;
  Specifications?: TrustedPartsPartSpecification[] | null;
  Distributors?: TrustedPartsDistributor[] | null;
};

type TrustedPartsApiResponse = {
  CurrentDate?: string;
  ResponseTime?: string;
  ErrorMessage?: string | null;
  Messages?: string[] | null;
  PartResults?: TrustedPartsPartResult[] | null;
};

type CacheEntry = {
  expiresAt: number;
  result: ProviderLookupResult;
};

const DEFAULT_BASE_URL = "https://api.trustedparts.com";
const SEARCH_PATH = "/v2/search";
const USER_AGENT = "XRayCounterApp/1.0";
const SOURCE_IP = "0.0.0.0";
const CACHE_TTL_MS = 15 * 60 * 1000;
const BATCH_SIZE = 50;
const BATCH_DELAY_MS = 200;
const HEALTHCHECK_PART = "BAV99";

export class TrustedPartsAdapter implements PricingAdapter {
  readonly name = "trustedparts";
  readonly mode = "live" as const;

  private static cache = new Map<string, CacheEntry>();

  private apiKey: string;
  private baseUrl: string;

  constructor(opts?: { apiKey?: string; baseUrl?: string }) {
    this.apiKey = (opts?.apiKey || process.env.TRUSTEDPARTS_API_KEY || "").trim();
    this.baseUrl = (opts?.baseUrl || process.env.TRUSTEDPARTS_BASE_URL || DEFAULT_BASE_URL)
      .trim()
      .replace(/\/$/, "");

    if (!this.apiKey) {
      throw new Error(
        "TrustedPartsAdapter: missing apiKey. Set TRUSTEDPARTS_API_KEY in env.",
      );
    }
  }

  async lookup(queries: ProviderQuery[]): Promise<ProviderLookupResult[]> {
    if (queries.length === 0) return [];

    const results = new Map<string, ProviderLookupResult>();
    const pending: ProviderQuery[] = [];

    for (const query of queries) {
      const cached = this.getCached(query.part_number);
      if (cached) {
        results.set(query.part_number, this.cloneResult(cached));
      } else {
        pending.push(query);
      }
    }

    for (let i = 0; i < pending.length; i += BATCH_SIZE) {
      const chunk = pending.slice(i, i + BATCH_SIZE);
      const chunkResults = await this.lookupChunk(chunk);
      for (const result of chunkResults) {
        results.set(result.part_number, this.cloneResult(result));
        this.setCached(result.part_number, result);
      }
      if (i + BATCH_SIZE < pending.length) {
        await delay(BATCH_DELAY_MS);
      }
    }

    return queries.map((query) => {
      const result = results.get(query.part_number);
      return result
        ? this.cloneResult(result)
        : {
            part_number: query.part_number,
            offers: [],
          };
    });
  }

  async health(): Promise<ProviderHealth> {
    const startedAt = Date.now();

    try {
      const response = await this.search(
        [
          {
            part_number: HEALTHCHECK_PART,
            qty: 1,
          },
        ],
        false,
      );

      return {
        ok: response.ok,
        provider: this.name,
        mode: this.mode,
        note: response.ok
          ? undefined
          : response.status === 401 || response.status === 403
            ? "unauthorized"
            : `http ${response.status}`,
        latency_ms: Date.now() - startedAt,
      };
    } catch (error) {
      return {
        ok: false,
        provider: this.name,
        mode: this.mode,
        note: (error as Error).message,
        latency_ms: Date.now() - startedAt,
      };
    }
  }

  private async lookupChunk(
    queries: ProviderQuery[],
  ): Promise<ProviderLookupResult[]> {
    const exactMatch = queries.length > 1;
    const response = await this.search(queries, exactMatch);

    if (response.status === 401 || response.status === 403) {
      console.warn(
        `[sourcing] TrustedParts authorization failed with status ${response.status}`,
      );
      return queries.map((query) => ({
        part_number: query.part_number,
        offers: [],
      }));
    }

    if (!response.ok) {
      throw new Error(
        `TrustedParts lookup failed with status ${response.status}: ${response.statusText}`,
      );
    }

    return this.mapResults(queries, response.body);
  }

  private async search(
    queries: ProviderQuery[],
    exactMatch: boolean,
  ): Promise<{
    ok: boolean;
    status: number;
    statusText: string;
    body: TrustedPartsApiResponse;
  }> {
    const firstQuery = queries[0];
    const payload: TrustedPartsApiRequest = {
      Queries: queries.map((query) => ({
        SearchToken: query.part_number,
      })),
      CountryCode: firstQuery?.region || "US",
      CurrencyCode: firstQuery?.currency || "USD",
      InStockOnly: false,
      ExactMatch: exactMatch,
      UseCachedData: false,
      LanguageCode: "en",
      IsCrawler: false,
      UserAgent: USER_AGENT,
      SourceIp: SOURCE_IP,
    };

    let response: Response;
    try {
      response = await fetch(`${this.baseUrl}${SEARCH_PATH}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${this.apiKey}`,
        },
        body: JSON.stringify(payload),
      });
    } catch (error) {
      throw new Error(
        `TrustedParts network error: ${(error as Error).message || "request failed"}`,
      );
    }

    let body: TrustedPartsApiResponse = {};
    const text = await response.text();
    if (text) {
      try {
        body = JSON.parse(text) as TrustedPartsApiResponse;
      } catch {
        body = {
          ErrorMessage: text,
        };
      }
    }

    return {
      ok: response.ok,
      status: response.status,
      statusText: response.statusText,
      body,
    };
  }

  private mapResults(
    queries: ProviderQuery[],
    body: TrustedPartsApiResponse,
  ): ProviderLookupResult[] {
    const parts = body.PartResults || [];

    return queries.map((query) => {
      const normalizedPartNumber = normalizePartNumber(query.part_number);
      const matches = parts.filter(
        (part) => normalizePartNumber(part.PartNumber) === normalizedPartNumber,
      );
      const primary = matches[0];
      const offers = matches.flatMap((part) => this.mapOffers(part, query.qty));
      const description = this.pickDescription(primary);
      const providerPartId = this.pickProviderPartId(primary);
      const alternates = this.collectAlternates(parts, normalizedPartNumber);

      return {
        part_number: query.part_number,
        offers,
        manufacturer: primary?.Manufacturer || undefined,
        description,
        provider_part_id: providerPartId,
        provider_alternates: alternates.length > 0 ? alternates : undefined,
      };
    });
  }

  private mapOffers(part: TrustedPartsPartResult, qty: number): Offer[] {
    const distributors = part.Distributors || [];
    const offers: Offer[] = [];

    for (const distributor of distributors) {
      const distributorName = distributor.Name?.trim();
      if (!distributorName) continue;

      for (const result of distributor.DistributorResults || []) {
        const unitPrice = this.pickUnitPrice(result.Pricing?.Prices || [], qty);
        if (unitPrice == null) continue;

        const stock = result.Stock?.QuantityOnHand ?? 0;
        const moq = this.pickMoq(result);
        offers.push({
          source: distributorName,
          unit_price: unitPrice,
          stock,
          lead_time_days: stock > 0 ? 3 : 14,
          moq,
        });
      }
    }

    return offers;
  }

  private pickUnitPrice(
    prices: TrustedPartsPrice[],
    qty: number,
  ): number | null {
    const validPrices = prices
      .filter(
        (price) =>
          typeof price.Quantity === "number" &&
          Number.isFinite(price.Quantity) &&
          typeof price.Amount === "number" &&
          Number.isFinite(price.Amount),
      )
      .sort((a, b) => a.Quantity - b.Quantity);

    if (validPrices.length === 0) return null;

    let selected = validPrices[0];
    for (const price of validPrices) {
      if (qty >= price.Quantity) {
        selected = price;
      } else {
        break;
      }
    }

    return selected.Amount;
  }

  private pickMoq(result: TrustedPartsDistributorResult): number {
    const pricingMoq = result.Pricing?.MinimumQuantity;
    if (typeof pricingMoq === "number" && Number.isFinite(pricingMoq) && pricingMoq > 0) {
      return pricingMoq;
    }

    for (const packaging of result.Packaging || []) {
      const packagingMoq = packaging.MinimumOrderQuantity;
      if (
        typeof packagingMoq === "number" &&
        Number.isFinite(packagingMoq) &&
        packagingMoq > 0
      ) {
        return packagingMoq;
      }
    }

    return 1;
  }

  private pickDescription(part?: TrustedPartsPartResult): string | undefined {
    if (!part) return undefined;
    for (const distributor of part.Distributors || []) {
      for (const result of distributor.DistributorResults || []) {
        const description = result.Description?.trim();
        if (description) return description;
      }
    }
    return undefined;
  }

  private pickProviderPartId(part?: TrustedPartsPartResult): string | undefined {
    if (!part) return undefined;
    for (const distributor of part.Distributors || []) {
      for (const result of distributor.DistributorResults || []) {
        const distributorPartNumber = result.DistributorPartNumber?.trim();
        if (distributorPartNumber) return distributorPartNumber;
      }
    }
    return part.PartNumber || undefined;
  }

  private collectAlternates(
    parts: TrustedPartsPartResult[],
    normalizedRequestedPart: string,
  ): string[] {
    const alternates = new Set<string>();
    for (const part of parts) {
      const candidate = part.PartNumber?.trim();
      if (!candidate) continue;
      if (normalizePartNumber(candidate) === normalizedRequestedPart) continue;
      alternates.add(candidate);
    }
    return [...alternates];
  }

  private getCached(partNumber: string): ProviderLookupResult | null {
    const cacheKey = normalizePartNumber(partNumber);
    const entry = TrustedPartsAdapter.cache.get(cacheKey);
    if (!entry) return null;
    if (entry.expiresAt <= Date.now()) {
      TrustedPartsAdapter.cache.delete(cacheKey);
      return null;
    }
    return entry.result;
  }

  private setCached(partNumber: string, result: ProviderLookupResult): void {
    TrustedPartsAdapter.cache.set(normalizePartNumber(partNumber), {
      expiresAt: Date.now() + CACHE_TTL_MS,
      result: this.cloneResult(result),
    });
  }

  private cloneResult(result: ProviderLookupResult): ProviderLookupResult {
    return {
      ...result,
      offers: result.offers.map((offer) => ({ ...offer })),
      provider_alternates: result.provider_alternates
        ? [...result.provider_alternates]
        : undefined,
    };
  }
}

function normalizePartNumber(partNumber?: string | null): string {
  return (partNumber || "").trim().toUpperCase();
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
