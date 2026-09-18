import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { reviseEbayQuantityBySku } from "@/lib/ebay/trading";

/**
 * TEMPORARY diagnostic-only endpoint — not linked from any UI. Lets us
 * call reviseEbayQuantityBySku directly and see its raw result (ack +
 * eBay's own error text) without digging through Vercel function logs,
 * while debugging why a quantity sync had no visible effect on a real
 * listing. Safe to delete once the sync is confirmed working: it
 * requires a signed-in session same as every other /api/parts route,
 * and only calls the same function app/api/parts/route.ts already
 * calls on every save.
 */
export async function GET(request: Request) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Sign in first." }, { status: 401 });
  }

  const { searchParams } = new URL(request.url);
  const sku = searchParams.get("sku");
  const qty = Number(searchParams.get("qty"));
  if (!sku || !Number.isFinite(qty)) {
    return NextResponse.json({ error: "Pass ?sku=...&qty=..." }, { status: 400 });
  }

  const result = await reviseEbayQuantityBySku(sku, qty);
  return NextResponse.json(result);
}
