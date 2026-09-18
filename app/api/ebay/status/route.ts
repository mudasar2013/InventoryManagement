import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { getStoredEbayToken } from "@/lib/ebay/token-store";

/**
 * Whether this shop's eBay account is currently connected (see
 * lib/ebay/oauth.ts) — never returns the tokens themselves, only
 * connection metadata, so this is safe to call from any signed-in
 * session without leaking a credential. Built as a quick way to
 * confirm a connect attempt actually landed in Redis; a real "eBay"
 * settings page (mirroring app/sources/page.tsx) can read from this
 * same shape later.
 */
export async function GET() {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Sign in first." }, { status: 401 });
  }

  const stored = await getStoredEbayToken();
  if (!stored) {
    return NextResponse.json({ connected: false });
  }

  return NextResponse.json({
    connected: true,
    connectedAt: stored.connectedAt,
    refreshTokenExpiresAt: stored.refreshTokenExpiresAt,
    hasCachedAccessToken: Boolean(stored.accessToken),
  });
}
