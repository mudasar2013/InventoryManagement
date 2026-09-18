"use client";

import { ExternalLink, RefreshCw, Sparkles } from "lucide-react";
import { useEffect, useState } from "react";

/**
 * The eBay panel on a part's detail page (see components/PartDetail.tsx)
 * — three independent, best-effort lookups against this part's SKU
 * (part_number), none of which block or affect the part's own data:
 *
 *  - Actual: what's really live on the eBay listing right now
 *    (app/api/ebay/description, GET half — lib/ebay/trading.ts).
 *  - Suggested: an AI-drafted description from this part's own catalog
 *    fields, generated on click (app/api/ebay/suggest-description —
 *    needs ANTHROPIC_API_KEY configured server-side).
 *  - Open-market pricing: what other sellers currently ask for this
 *    part number on eBay (app/api/ebay/market-price — lib/ebay/browse.ts).
 *
 * "New" pricing from Marcone/Tribles/Encompass isn't built yet — see
 * the note at the bottom of this panel; automating logins to those
 * dealer sites needs a separate decision before it's built (ToS/
 * account-risk considerations), so this panel is honest about it
 * rather than silently omitting the row.
 */

interface EbayListing {
  itemId: string;
  title: string | null;
  descriptionHtml: string | null;
}

interface MarketPriceResult {
  query: string;
  listings: {
    title: string;
    price: number | null;
    currency: string | null;
    condition: string | null;
    itemWebUrl: string | null;
    seller: string | null;
  }[];
  priceStats: { min: number; max: number; average: number; count: number } | null;
}

export function EbayPanel({
  partNumber,
  description,
  category,
}: {
  partNumber: string;
  description?: string;
  category?: string;
}) {
  const [listing, setListing] = useState<EbayListing | null>(null);
  const [listingNote, setListingNote] = useState<string | null>(null);
  const [listingLoading, setListingLoading] = useState(true);

  const [market, setMarket] = useState<MarketPriceResult | null>(null);
  const [marketError, setMarketError] = useState<string | null>(null);
  const [marketLoading, setMarketLoading] = useState(true);

  const [suggested, setSuggested] = useState<string | null>(null);
  const [suggesting, setSuggesting] = useState(false);
  const [suggestError, setSuggestError] = useState<string | null>(null);

  // Wrapped in their own async functions (rather than setState calls
  // sitting directly in the effect body) purely to satisfy the
  // set-state-in-effect lint rule — see the same pattern in
  // components/AddPartForm.tsx's UPN# preview effect.
  useEffect(() => {
    let cancelled = false;

    async function loadListing() {
      setListingLoading(true);
      setListing(null);
      setListingNote(null);
      try {
        const response = await fetch(`/api/ebay/description?sku=${encodeURIComponent(partNumber)}`);
        const payload = await response.json();
        if (cancelled) return;
        if (payload.listing) {
          setListing(payload.listing);
        } else {
          setListingNote(payload.note ?? payload.error ?? "Not listed on eBay.");
        }
      } catch {
        if (!cancelled) setListingNote("Couldn't reach eBay.");
      } finally {
        if (!cancelled) setListingLoading(false);
      }
    }

    async function loadMarketPrice() {
      setMarketLoading(true);
      setMarket(null);
      setMarketError(null);
      try {
        const response = await fetch(`/api/ebay/market-price?q=${encodeURIComponent(partNumber)}`);
        const payload = await response.json();
        if (cancelled) return;
        if (payload.result) {
          setMarket(payload.result);
        } else {
          setMarketError(payload.error ?? "Couldn't check market pricing.");
        }
      } catch {
        if (!cancelled) setMarketError("Couldn't reach eBay.");
      } finally {
        if (!cancelled) setMarketLoading(false);
      }
    }

    loadListing();
    loadMarketPrice();
    return () => {
      cancelled = true;
    };
  }, [partNumber]);

  async function generateSuggestion() {
    setSuggesting(true);
    setSuggestError(null);
    try {
      const response = await fetch("/api/ebay/suggest-description", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ part_number: partNumber, description, category }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to generate a description.");
      }
      setSuggested(payload.suggested);
    } catch (err) {
      setSuggestError(err instanceof Error ? err.message : "Failed to generate a description.");
    } finally {
      setSuggesting(false);
    }
  }

  return (
    <section className="space-y-3">
      <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">eBay</h2>

      <div className="space-y-3 rounded-2xl border border-stone-200 bg-white p-4">
        <div>
          <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
            Actual description
          </p>
          {listingLoading ? (
            <p className="mt-1 text-sm text-stone-400">Checking eBay…</p>
          ) : listing ? (
            <div className="mt-1 space-y-1">
              <a
                href={`https://www.ebay.com/itm/${listing.itemId}`}
                target="_blank"
                rel="noreferrer"
                className="inline-flex items-center gap-1 text-sm font-semibold text-sky-700 hover:underline"
              >
                {listing.title ?? "View listing"}
                <ExternalLink className="size-3.5" />
              </a>
              {listing.descriptionHtml ? (
                <div
                  className="prose prose-sm max-w-none text-stone-700"
                  dangerouslySetInnerHTML={{ __html: listing.descriptionHtml }}
                />
              ) : (
                <p className="text-sm text-stone-400">No description text on the listing.</p>
              )}
            </div>
          ) : (
            <p className="mt-1 text-sm text-stone-400">{listingNote}</p>
          )}
        </div>

        <div className="border-t border-stone-100 pt-3">
          <div className="flex items-center justify-between gap-2">
            <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
              Suggested description
            </p>
            <button
              type="button"
              onClick={generateSuggestion}
              disabled={suggesting}
              className="inline-flex items-center gap-1.5 rounded-full bg-amber-50 px-3 py-1.5 text-xs font-semibold text-amber-800 hover:bg-amber-100 disabled:opacity-60"
            >
              {suggesting ? (
                <RefreshCw className="size-3.5 animate-spin" />
              ) : (
                <Sparkles className="size-3.5" />
              )}
              {suggesting ? "Generating…" : suggested ? "Regenerate" : "Generate"}
            </button>
          </div>
          {suggestError ? (
            <p className="mt-1 text-sm text-rose-700">{suggestError}</p>
          ) : suggested ? (
            <p className="mt-1 whitespace-pre-wrap text-sm text-stone-700">{suggested}</p>
          ) : (
            <p className="mt-1 text-sm text-stone-400">Not generated yet.</p>
          )}
        </div>

        <div className="border-t border-stone-100 pt-3">
          <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
            Pricing — open market (eBay)
          </p>
          {marketLoading ? (
            <p className="mt-1 text-sm text-stone-400">Checking eBay…</p>
          ) : marketError ? (
            <p className="mt-1 text-sm text-rose-700">{marketError}</p>
          ) : market?.priceStats ? (
            <p className="mt-1 text-sm text-stone-700">
              ${market.priceStats.min.toFixed(2)} – ${market.priceStats.max.toFixed(2)}, average $
              {market.priceStats.average.toFixed(2)} across {market.priceStats.count} listing
              {market.priceStats.count === 1 ? "" : "s"}
            </p>
          ) : (
            <p className="mt-1 text-sm text-stone-400">No comparable active listings found.</p>
          )}
          <p className="mt-2 text-xs text-stone-400">
            Pricing — new (Marcone / Tribles / Encompass): not built yet — those sites need a
            dealer login with no public pricing API, so this needs a decision on how to safely
            automate that before it&apos;s added here.
          </p>
        </div>
      </div>
    </section>
  );
}
