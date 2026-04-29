import { NextRequest, NextResponse } from "next/server";
import { ADMIN_COOKIE } from "@/lib/admin";

export async function POST(req: NextRequest) {
  try {
    const configured = process.env.ADMIN_TOKEN;
    if (!configured) {
      return NextResponse.json({ ok: true, authDisabled: true });
    }

    const body = await req.json().catch(() => ({}));
    const token = String(body?.token || "").trim();
    if (!token || token !== configured) {
      return NextResponse.json({ error: "Invalid admin token" }, { status: 401 });
    }

    const res = NextResponse.json({ ok: true });
    res.cookies.set(ADMIN_COOKIE, token, {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.NODE_ENV === "production",
      path: "/",
      maxAge: 60 * 60 * 24 * 30,
    });
    return res;
  } catch (e: unknown) {
    const message = e instanceof Error ? e.message : "Login failed";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
