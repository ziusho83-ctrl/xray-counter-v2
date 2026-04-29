import { NextRequest, NextResponse } from "next/server";
import { isAdmin } from "@/lib/admin";

export async function GET(req: NextRequest) {
  const requiresAuth = !!process.env.ADMIN_TOKEN;
  const authed = isAdmin(req);
  return NextResponse.json({ ok: true, requiresAuth, authed });
}
