import { readEbayConfig } from "./config";

/**
 * eBay's Buy Browse API — used only for "open market" pricing (see
 * app/api/ebay/market-price/route.ts): searching active listings by
 * keyword (a part number) to see what other sellers currently ask for
 * the same part. This is a read-only public search, so it uses an
 * eBay *application* access token (client_credentials grant) rather
 * than this shop's own user token from lib/ebay/oauth.ts — it isn't
 * acting on this shop's account at all, just reading public listing
 * data, and application tokens don't need a connected seller account
 * to exist.
 */

const BROWSE_SCOPE = "https://api.ebay.com/oauth/api_scope/buy.browse";

function ebayIdentityTokenUrl(env: "production" | "sandbox"): string {
  return env === "sandbox"
    ? "https://api.sandbox.ebay.com/identity/v1/oauth2/token"
    : "https://api.ebay.com/identity/v1/oauth2/token";
}

function ebayBrowseApiUrl(env: "production" | "sandbox"): string {
  return env === "sandbox"
    ? "https://api.sandbox.ebay.com/buy/browse/v1/item_summary/search"
    : "https://api.ebay.com/buy/browse/v1/item_summary/search";
}

async function getApplicationAccessToken(): Promise<
  { token: string } | { error: string }
> {
  const config = readEbayConfig();
  if (!config) {
    return { error: "eBay app credentials are not configured." };
  }
  const basicAuth = `Basic ${Buffer.from(`${config.appId}:${config.certId}`).toString("base64")}`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    scope: BROWSE_SCOPE,
  });

  let response: Response;
  try {
    response = await fetch(ebayIdentityTokenUrl(config.env), {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Authorization: basicAuth,
      },
      body,
    });
  } catch (error) {
    return { error: error instanceof Error ? error.message : "eBay token request failed." };
  }

  const payload = (await response.json().catch(() => ({}))) as {
    access_token?: string;
    error?: string;
    error_description?: string;
  };
  if (!response.ok || !payload.access_token) {
    return {
      error: `eBay application token request failed: ${payload.error ?? response.status} ${payload.error_description ?? ""}`.trim(),
    };
  }
  return { token: payload.access_token };
}

export interface CompetitorListing {
  title: string;
  price: number | null;
  currency: string | null;
  condition: string | null;
  itemWebUrl: string | null;
  seller: string | null;
}

export interface MarketPriceSummary {
  query: string;
  listings: CompetitorListing[];
  priceStats: { min: number; max: number; average: number; count: number } | null;
}

/**
 * Searches eBay's active listings for a keyword (normally a part
 * number) and summarizes what other sellers currently ask for it —
 * the "Open market" pricing point for a part (see
 * components/PartDetail.tsx). Never throws: returns an `error`
 * instead, since this is a supplementary lookup that shouldn't break
 * the page it's shown on.
 */
export async function searchEbayMarketPrice(
  query: string,
  limit = 20,
): Promise<{ result: MarketPriceSummary | null; error: string | null }> {
  const config = readEbayConfig();
  if (!config) {
    return { result: null, error: "eBay app credentials are not configured." };
  }

  const tokenResult = await getApplicationAccessToken();
  if ("error" in tokenResult) {
    return { result: null, error: tokenResult.error };
  }

  const url = new URL(ebayBrowseApiUrl(config.env));
  url.searchParams.set("q", query);
  url.searchParams.set("limit", String(Math.min(Math.max(limit, 1), 50)));

  let response: Response;
  try {
    response = await fetch(url.toString(), {
      headers: {
        Authorization: `Bearer ${tokenResult.token}`,
        "X-EBAY-C-MARKETPLACE-ID": "EBAY_US",
        "Content-Type": "application/json",
      },
    });
  } catch (error) {
    return { result: null, error: error instanceof Error ? error.message : "eBay Browse API request failed." };
  }

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message =
      Array.isArray(payload?.errors) && payload.errors[0]?.message
        ? payload.errors[0].message
        : `Browse API returned ${response.status}`;
    return { result: null, error: message };
  }

  const rawItems: unknown[] = Array.isArray(payload?.itemSummaries) ? payload.itemSummaries : [];
  const listings: CompetitorListing[] = rawItems.map((raw) => {
    const item = raw as Record<string, unknown>;
    const price = item.price as Record<string, unknown> | undefined;
    const seller = item.seller as Record<string, unknown> | undefined;
    const priceValue = price?.value ? Number(price.value) : null;
    return {
      title: typeof item.title === "string" ? item.title : "",
      price: Number.isFinite(priceValue) ? priceValue : null,
      currency: typeof price?.currency === "string" ? (price.currency as string) : null,
      condition: typeof item.condition === "string" ? (item.condition as string) : null,
      itemWebUrl: typeof item.itemWebUrl === "string" ? (item.itemWebUrl as string) : null,
      seller: typeof seller?.username === "string" ? (seller.username as string) : null,
    };
  });

  const prices = listings.map((l) => l.price).filter((p): p is number => p !== null);
  const priceStats = prices.length
    ? {
        min: Math.min(...prices),
        max: Math.max(...prices),
        average: Math.round((prices.reduce((a, b) => a + b, 0) / prices.length) * 100) / 100,
        count: prices.length,
      }
    : null;

  return { result: { query, listings, priceStats }, error: null };
}
