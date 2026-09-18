import { ebayTradingApiUrl, readEbayConfig } from "./config";
import { getValidAccessToken } from "./oauth";

/**
 * eBay's Trading API (XML, `api.dll`) — used to sync a part's quantity
 * to its live eBay listing by SKU (see app/api/parts/route.ts, which
 * calls reviseEbayQuantityBySku after every Add/Edit Part save).
 *
 * This does NOT use the newer Sell Inventory API
 * (bulkUpdatePriceQuantity / getOffers): that API's SKU-based
 * endpoints only recognize listings that were created (or migrated)
 * through the Inventory API itself, and 404 on ordinary listings made
 * through Seller Hub — which is how this shop's listings exist today.
 *
 * It also does NOT call ReviseInventoryStatus with a bare <SKU> the
 * way an earlier version of this file did — that only works when the
 * listing's Item.InventoryTrackingMethod was set to SKU at the time
 * it was originally listed via the Trading API, which Seller-Hub-
 * created listings generally aren't, even when their Custom Label
 * (SKU) field is set and matches exactly (confirmed against a real
 * listing: eBay returned "Invalid SKU number" despite the SKU being
 * correct — see https://developer.ebay.com/support/kb-article?KBid=1465).
 * Every call here instead resolves the listing's ItemID from its SKU
 * first (via GetSellerList's SKUArray filter, which eBay's own docs
 * confirm works "regardless of Item.InventoryTrackingMethod"), then
 * acts on that ItemID, which always works.
 *
 * Ending a listing also can't be done by setting Quantity to 0:
 * ReviseInventoryStatus rejects that outright ("Invalid quantity...
 * must be greater than 0 for active items" — same KB article above).
 * EndFixedPriceItem is the real "end this listing" call, and it also
 * needs an ItemID rather than a SKU.
 */

const COMPATIBILITY_LEVEL = "1155";
const SITE_ID = "0"; // eBay US
const END_TIME_WINDOW_DAYS = 30;

interface RawTradingApiResult {
  ok: boolean;
  ack: string | null;
  errors: string[];
  raw: string;
}

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

async function callTradingApi(callName: string, bodyXml: string): Promise<RawTradingApiResult> {
  const config = readEbayConfig();
  if (!config) {
    return { ok: false, ack: null, errors: ["eBay app credentials are not configured."], raw: "" };
  }

  let accessToken: string | null;
  try {
    accessToken = await getValidAccessToken();
  } catch (error) {
    return {
      ok: false,
      ack: null,
      errors: [error instanceof Error ? error.message : "Failed to get an eBay access token."],
      raw: "",
    };
  }
  if (!accessToken) {
    return { ok: false, ack: null, errors: ["eBay account is not connected."], raw: "" };
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
      raw: "",
    };
  }

  const text = await response.text();
  const ack = extractTag(text, "Ack");
  const errors = extractErrorMessages(text);
  return { ok: ack === "Success" || ack === "Warning", ack, errors, raw: text };
}

/**
 * Resolves a SKU (the part's `part_number`) to the ItemID of the live
 * eBay listing using that Custom Label — see the module doc comment
 * for why this lookup is required before every revise/end call. Scoped
 * with EndTimeFrom/EndTimeTo (now .. +30 days) rather than
 * StartTimeFrom/To, which is eBay's own documented way to reliably
 * catch Good-Til-Cancelled listings regardless of how long ago they
 * were first listed — see
 * https://developer.ebay.com/support/kb-article?KBid=5020. Returns
 * null (not an error) when no listing has this SKU, which is the
 * common case for parts that simply aren't on eBay.
 */
async function findItemIdBySku(sku: string): Promise<{ itemId: string | null; errors: string[] }> {
  const now = new Date();
  const windowEnd = new Date(now.getTime() + END_TIME_WINDOW_DAYS * 24 * 60 * 60 * 1000);
  const body = `<?xml version="1.0" encoding="utf-8"?>
<GetSellerListRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <SKUArray>
    <SKU>${xmlEscape(sku)}</SKU>
  </SKUArray>
  <EndTimeFrom>${now.toISOString()}</EndTimeFrom>
  <EndTimeTo>${windowEnd.toISOString()}</EndTimeTo>
  <Pagination>
    <EntriesPerPage>10</EntriesPerPage>
    <PageNumber>1</PageNumber>
  </Pagination>
</GetSellerListRequest>`;

  const result = await callTradingApi("GetSellerList", body);
  if (!result.ok) {
    return { itemId: null, errors: result.errors.length ? result.errors : [`GetSellerList: ${result.ack ?? "failed"}`] };
  }
  const itemId = extractTag(result.raw, "ItemID");
  return { itemId, errors: [] };
}

/**
 * Sets a live eBay listing's available quantity by SKU (resolved to
 * ItemID first — see findItemIdBySku), or ends the listing when
 * quantity is 0 or less. `sku` is expected to be the part's
 * `part_number` — see app/api/parts/route.ts, which calls this after
 * every successful quantity save. Never throws: callers get back a
 * result they can log or surface, so a part's own save never fails
 * just because eBay rejected the sync or this SKU isn't listed there
 * at all (the common case for most parts).
 */
export async function reviseEbayQuantityBySku(
  sku: string,
  quantity: number,
): Promise<TradingApiResult> {
  const { itemId, errors: lookupErrors } = await findItemIdBySku(sku);
  if (!itemId) {
    return {
      ok: false,
      ack: null,
      errors: lookupErrors.length ? lookupErrors : [`No eBay listing found with SKU "${sku}".`],
    };
  }

  const safeQuantity = Math.max(0, Math.trunc(quantity));
  if (safeQuantity > 0) {
    const body = `<?xml version="1.0" encoding="utf-8"?>
<ReviseInventoryStatusRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <InventoryStatus>
    <ItemID>${xmlEscape(itemId)}</ItemID>
    <Quantity>${safeQuantity}</Quantity>
  </InventoryStatus>
</ReviseInventoryStatusRequest>`;
    const result = await callTradingApi("ReviseInventoryStatus", body);
    return { ok: result.ok, ack: result.ack, errors: result.errors };
  }

  const body = `<?xml version="1.0" encoding="utf-8"?>
<EndFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <ItemID>${xmlEscape(itemId)}</ItemID>
  <EndingReason>NotAvailable</EndingReason>
</EndFixedPriceItemRequest>`;
  const result = await callTradingApi("EndFixedPriceItem", body);
  return { ok: result.ok, ack: result.ack, errors: result.errors };
}

function xmlUnescape(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

export interface EbayListingSnapshot {
  itemId: string;
  title: string | null;
  descriptionHtml: string | null;
}

/**
 * Fetches a live eBay listing's current title and description by SKU
 * (via the same findItemIdBySku lookup every other call here uses) —
 * the "Actual" side of the Add/Edit Part eBay-description panel (see
 * components/PartDetail.tsx), compared against an AI-drafted
 * "Suggested" description from app/api/ebay/suggest-description.
 * Returns null when this SKU isn't listed on eBay at all, same as
 * reviseEbayQuantityBySku's "not found" case.
 */
export async function getEbayListingBySku(sku: string): Promise<{
  listing: EbayListingSnapshot | null;
  errors: string[];
}> {
  const { itemId, errors: lookupErrors } = await findItemIdBySku(sku);
  if (!itemId) {
    return { listing: null, errors: lookupErrors };
  }

  const body = `<?xml version="1.0" encoding="utf-8"?>
<GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <ItemID>${xmlEscape(itemId)}</ItemID>
  <DetailLevel>ItemReturnDescription</DetailLevel>
</GetItemRequest>`;
  const result = await callTradingApi("GetItem", body);
  if (!result.ok) {
    return { listing: null, errors: result.errors.length ? result.errors : [`GetItem: ${result.ack ?? "failed"}`] };
  }

  const title = extractTag(result.raw, "Title");
  const rawDescription = extractTag(result.raw, "Description");
  return {
    listing: {
      itemId,
      title: title ? xmlUnescape(title) : null,
      descriptionHtml: rawDescription ? xmlUnescape(rawDescription) : null,
    },
    errors: [],
  };
}

