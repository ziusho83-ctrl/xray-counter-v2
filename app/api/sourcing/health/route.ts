import { NextResponse } from "next/server";
import { getPricingAdapter } from "@/lib/sourcing/adapters";

export const runtime = "nodejs";

export async function GET() {
  const adapter = getPricingAdapter();
  const health = await adapter.health();
  return NextResponse.json({
    sourcing_version: "0.3",
    provider: health,
    env: {
      SOURCING_MODE: process.env.SOURCING_MODE || "(auto)",
      TRUSTEDPARTS_API_KEY: process.env.TRUSTEDPARTS_API_KEY ? "set" : "unset",
      TRUSTEDPARTS_BASE_URL: process.env.TRUSTEDPARTS_BASE_URL ? "set" : "unset",
    },
  });
}
