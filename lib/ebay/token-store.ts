import { getRedis, isRedisConfigured } from "../redis";

/**
 * Stores the one eBay user (refresh) token this shop's Production
 * keyset is authorized under — see lib/ebay/oauth.ts, which exchanges
 * the OAuth consent code for this and keeps a short-lived access token
 * minted from it. This app manages exactly one eBay seller account
 * (this shop's own), so — unlike sharepoint-source-store.ts, which
 * holds a list — there is only ever one record here, at a fixed key.
 *
 * Same Redis-or-nothing philosophy as the other app-managed stores:
 * with no Redis configured, eBay features simply report "not
 * connected" rather than erroring.
 */
export interface StoredEbayToken {
  refreshToken: string;
  /** Seconds from grant until refreshToken itself expires (eBay's
   *  `refresh_token_expires_in`, ~18 months) — surfaced on the eBay
   *  status page so a shop owner sees "reconnect by <date>" well
   *  before it silently stops working, rather than only finding out
   *  when a decrement/end-listing call starts failing. */
  refreshTokenExpiresAt: string;
  /** Cached short-lived access token (eBay's `access_token`, ~2 hours)
   *  so every API call doesn't need its own refresh round-trip —
   *  see getValidAccessToken. */
  accessToken?: string;
  accessTokenExpiresAt?: string;
  connectedAt: string;
}

const STORE_KEY = "inventory:ebay-token";

export const isEbayTokenStoreConfigured = isRedisConfigured;

export async function getStoredEbayToken(): Promise<StoredEbayToken | null> {
  const redis = getRedis();
  if (!redis) {
    return null;
  }
  return (await redis.get<StoredEbayToken>(STORE_KEY)) ?? null;
}

export async function setStoredEbayToken(token: StoredEbayToken): Promise<void> {
  const redis = getRedis();
  if (!redis) {
    throw new Error("Redis is not configured — cannot store the eBay token.");
  }
  await redis.set(STORE_KEY, token);
}

/** Updates just the cached access token on an already-connected
 *  record, leaving the refresh token and connectedAt untouched — used
 *  by getValidAccessToken after a refresh. Does nothing (rather than
 *  throwing) if this shop was never connected in the first place;
 *  callers only reach here after confirming a refresh token exists. */
export async function updateStoredAccessToken(
  accessToken: string,
  accessTokenExpiresAt: string,
): Promise<void> {
  const redis = getRedis();
  if (!redis) {
    return;
  }
  const existing = await redis.get<StoredEbayToken>(STORE_KEY);
  if (!existing) {
    return;
  }
  await redis.set(STORE_KEY, { ...existing, accessToken, accessTokenExpiresAt });
}

export async function clearStoredEbayToken(): Promise<void> {
  const redis = getRedis();
  if (!redis) {
    return;
  }
  await redis.del(STORE_KEY);
}
