import { NextRequest } from "next/server";

export const ADMIN_COOKIE = "xray_admin_token";

export function isAdmin(req: NextRequest): boolean {
  const configured = process.env.ADMIN_TOKEN;
  // Optional security gate: enforce only when ADMIN_TOKEN is configured.
  if (!configured) return true;

  const headerToken = req.headers.get("x-admin-token") || "";
  const cookieToken = req.cookies.get(ADMIN_COOKIE)?.value || "";
  const provided = headerToken || cookieToken;
  return provided === configured;
}
