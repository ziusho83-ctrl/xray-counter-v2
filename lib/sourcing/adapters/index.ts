import type { PricingAdapter } from "./types";
import { MockPricingAdapter } from "./mock";
import { TrustedPartsAdapter } from "./trustedparts";

/**
 * Orchestrates which pricing adapter is active.
 *
 * Priority:
 *   1. Explicit override via SOURCING_MODE env var:
 *        SOURCING_MODE=mock           -> always mock
 *        SOURCING_MODE=trustedparts   -> always TrustedParts (fail if unconfigured)
 *   2. Auto-detect: if TRUSTEDPARTS_API_KEY + TRUSTEDPARTS_BASE_URL are set,
 *      use TrustedParts. Otherwise mock.
 *
 * The resolver is intentionally synchronous and cheap so it can be called
 * inside request handlers without overhead. It memoizes the instance per
 * process (adapters are stateless aside from config).
 */

let cached: { key: string; adapter: PricingAdapter } | null = null;

function resolveMode(): "mock" | "trustedparts" {
  const override = (process.env.SOURCING_MODE || "").trim().toLowerCase();
  if (override === "mock") return "mock";
  if (override === "trustedparts" || override === "live") return "trustedparts";

  const hasKey = !!process.env.TRUSTEDPARTS_API_KEY;
  const hasUrl = !!process.env.TRUSTEDPARTS_BASE_URL;
  if (hasKey && hasUrl) return "trustedparts";
  return "mock";
}

function buildAdapter(mode: "mock" | "trustedparts"): PricingAdapter {
  if (mode === "trustedparts") {
    try {
      return new TrustedPartsAdapter({
        apiKey: process.env.TRUSTEDPARTS_API_KEY || "",
        baseUrl: process.env.TRUSTEDPARTS_BASE_URL || "",
      });
    } catch (e) {
      // Fail-open: if TrustedParts is misconfigured, fall back to mock rather
      // than breaking the UI. The /api/sourcing/health endpoint surfaces the
      // real state so we can see what's going on.
      console.warn(
        "[sourcing] TrustedParts adapter init failed, falling back to mock:",
        (e as Error).message,
      );
      return new MockPricingAdapter();
    }
  }
  return new MockPricingAdapter();
}

export function getPricingAdapter(): PricingAdapter {
  const mode = resolveMode();
  const key = `${mode}:${process.env.TRUSTEDPARTS_BASE_URL || ""}`;
  if (cached && cached.key === key) return cached.adapter;
  const adapter = buildAdapter(mode);
  cached = { key, adapter };
  return adapter;
}

export function clearAdapterCache(): void {
  cached = null;
}
