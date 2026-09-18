import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { searchEbayMarketPrice } from "@/lib/ebay/browse";

/**
 * The "Open market" pricing point for a part (see
 * components/PartDetail.tsx) — what other sellers currently ask for
 * the same part number on eBay right now, via the public Browse API
 * (see lib/ebay/browse.ts). `q` is normally the part's part_number.
 */
export async function GET(request: Request) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Sign in first." }, { status: 401 });
  }

  const q = new URL(request.url).searchParams.get("q");
  if (!q) {
    return NextResponse.json({ error: "q is required." }, { status: 400 });
  }

  const { result, error } = await searchEbayMarketPrice(q);
  if (error) {
    return NextResponse.json({ error }, { status: 502 });
  }
  return NextResponse.json({ result });
}
