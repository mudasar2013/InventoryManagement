import { ebayTradingApiUrl, readEbayConfig } from "./config";
import { getValidAccessToken } from "./oauth";

/**
 * eBay's Trading API (XML, `api.dll`) — used only for
 * ReviseInventoryStatus, which is the one call that can set a live
 * listing's quantity straight from a SKU with no ItemID lookup at all.
 * This deliberately does NOT use the newer Sell Inventory API
 * (bulkUpdatePriceQuantity / getOffers): that API's SKU-based
 * endpoints only recognize listings that were created (or migrated)
 * through the Inventory API itself, and 404 on ordinary listings made
 * through Seller Hub — which is how this shop's listings exist today.
 * ReviseInventoryStatus works on those directly.
 *
 * Setting Quantity to 0 is deliberately how "end this listing" is
 * implemented: eBay's default behavior (the "Out-of-Stock" listing
 * feature turned off, which is the default) ends a fixed-price
 * listing automatically once its available quantity hits zero — no
 * separate EndItem/ItemID call needed. If this shop ever turns the
 * Out-of-Stock feature ON in its eBay account, a 0 quantity will
 * instead just mark the listing unavailable rather than ending it;
 * see https://developer.ebay.com/api-docs/user-guides/static/trading-user-guide/out-of-stock-operation.html.
 */

const COMPATIBILITY_LEVEL = "1155";
const SITE_ID = "0"; // eBay US

export interface TradingApiResult {
  ok: boolean;
  ack: string | null;
  errors: string[];
}

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function extractTag(xml: string, tag: string): string | null {
  const match = xml.match(new RegExp(`<${tag}>([^<]*)</${tag}>`));
  return match ? match[1] : null;
}

function extractErrorMessages(xml: string): string[] {
  const messages: string[] = [];
  const regex = /<ShortMessage>([^<]*)<\/ShortMessage>/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(xml)) !== null) {
    messages.push(match[1]);
  }
  return messages;
}

async function callTradingApi(callName: string, bodyXml: string): Promise<TradingApiResult> {
  const config = readEbayConfig();
  if (!config) {
    return { ok: false, ack: null, errors: ["eBay app credentials are not configured."] };
  }

  let accessToken: string | null;
  try {
    accessToken = await getValidAccessToken();
  } catch (error) {
    return {
      ok: false,
      ack: null,
      errors: [error instanceof Error ? error.message : "Failed to get an eBay access token."],
    };
  }
  if (!accessToken) {
    return { ok: false, ack: null, errors: ["eBay account is not connected."] };
  }

  let response: Response;
  try {
    response = await fetch(ebayTradingApiUrl(config.env), {
      method: "POST",
      headers: {
        "Content-Type": "text/xml",
        "X-EBAY-API-CALL-NAME": callName,
        "X-EBAY-API-SITEID": SITE_ID,
        "X-EBAY-API-COMPATIBILITY-LEVEL": COMPATIBILITY_LEVEL,
        "X-EBAY-API-IAF-TOKEN": accessToken,
        "X-EBAY-API-APP-NAME": config.appId,
        "X-EBAY-API-DEV-NAME": config.devId,
        "X-EBAY-API-CERT-NAME": config.certId,
      },
      body: bodyXml,
    });
  } catch (error) {
    return {
      ok: false,
      ack: null,
      errors: [error instanceof Error ? error.message : "eBay Trading API request failed."],
    };
  }

  const text = await response.text();
  const ack = extractTag(text, "Ack");
  const errors = extractErrorMessages(text);
  return { ok: ack === "Success" || ack === "Warning", ack, errors };
}

/**
 * Sets a live eBay listing's available quantity by SKU. Pass 0 to end
 * the listing (see the module doc comment above for why that's
 * sufficient). `sku` is expected to be the part's `part_number` — see
 * app/api/parts/route.ts, which calls this after every successful
 * quantity save. Never throws: callers get back a result they can log
 * or surface, so a part's own save never fails just because eBay
 * rejected or couldn't reach the sync (e.g. this SKU isn't actually
 * listed on eBay, which is the common case for most parts).
 */
export async function reviseEbayQuantityBySku(
  sku: string,
  quantity: number,
): Promise<TradingApiResult> {
  const safeQuantity = Math.max(0, Math.trunc(quantity));
  const body = `<?xml version="1.0" encoding="utf-8"?>
<ReviseInventoryStatusRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <InventoryStatus>
    <SKU>${xmlEscape(sku)}</SKU>
    <Quantity>${safeQuantity}</Quantity>
  </InventoryStatus>
</ReviseInventoryStatusRequest>`;
  return callTradingApi("ReviseInventoryStatus", body);
}
