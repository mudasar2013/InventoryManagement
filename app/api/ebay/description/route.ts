import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { getEbayListingBySku } from "@/lib/ebay/trading";

/**
 * The "Actual" half of the eBay description panel (see
 * components/PartDetail.tsx) — what's really live on eBay right now
 * for this part's SKU, straight from the listing itself. Pairs with
 * app/api/ebay/suggest-description, the AI-drafted "Suggested" half.
 */
export async function GET(request: Request) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Sign in first." }, { status: 401 });
  }

  const sku = new URL(request.url).searchParams.get("sku");
  if (!sku) {
    return NextResponse.json({ error: "sku is required." }, { status: 400 });
  }

  const { listing, errors } = await getEbayListingBySku(sku);
  if (!listing) {
    return NextResponse.json({ listing: null, note: errors.join("; ") || "Not listed on eBay." });
  }
  return NextResponse.json({ listing });
}
