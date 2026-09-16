import { ebayOAuthTokenUrl, readEbayConfig } from "./config";
import {
  getStoredEbayToken,
  setStoredEbayToken,
  updateStoredAccessToken,
} from "./token-store";

/**
 * eBay's OAuth user-token flow — see
 * https://developer.ebay.com/api-docs/static/oauth-tokens.html.
 *
 * This app never sees the seller's eBay password: the shop owner signs
 * in on eBay's own page (today, via the "Sign in to Production for
 * OAuth" button on the eBay Developer Portal itself — this app has no
 * "Connect eBay" button of its own yet), eBay redirects back to
 * app/api/ebay/oauth/callback with a short-lived `code`, and
 * exchangeCodeForTokens below trades that for an access token (~2
 * hours) and a refresh token (~18 months) — see lib/ebay/token-store.ts
 * for where those get kept. Every other eBay API call should go
 * through getValidAccessToken() rather than reading the stored access
 * token directly, so a call never has to think about whether it's
 * about to expire.
 */

interface EbayTokenResponse {
  access_token: string;
  expires_in: number;
  refresh_token?: string;
  refresh_token_expires_in?: number;
  token_type: string;
  error?: string;
  error_description?: string;
}

function basicAuthHeader(appId: string, certId: string): string {
  return `Basic ${Buffer.from(`${appId}:${certId}`).toString("base64")}`;
}

async function requestToken(body: URLSearchParams): Promise<EbayTokenResponse> {
  const config = readEbayConfig();
  if (!config) {
    throw new Error(
      "EBAY_APP_ID/EBAY_CERT_ID/EBAY_DEV_ID/EBAY_RU_NAME are not fully configured.",
    );
  }
  const response = await fetch(ebayOAuthTokenUrl(config.env), {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Authorization: basicAuthHeader(config.appId, config.certId),
    },
    body,
  });
  const payload = (await response.json().catch(() => ({}))) as EbayTokenResponse;
  if (!response.ok || payload.error) {
    throw new Error(
      `eBay token request failed: ${payload.error ?? response.status} ${payload.error_description ?? ""}`.trim(),
    );
  }
  return payload;
}

/**
 * Trades the one-time authorization `code` eBay's redirect carried
 * (valid ~5 minutes — see app/api/ebay/oauth/callback/route.ts) for a
 * real access + refresh token pair, and stores both. Called exactly
 * once per "connect" (or "reconnect after the refresh token expires,
 * ~18 months out") — every ordinary API call afterward uses
 * getValidAccessToken() instead, never this.
 */
export async function exchangeCodeForTokens(code: string): Promise<void> {
  const config = readEbayConfig();
  if (!config) {
    throw new Error(
      "EBAY_APP_ID/EBAY_CERT_ID/EBAY_DEV_ID/EBAY_RU_NAME are not fully configured.",
    );
  }
  const body = new URLSearchParams({
    grant_type: "authorization_code",
    code,
    redirect_uri: config.ruName,
  });
  const payload = await requestToken(body);
  if (!payload.refresh_token || !payload.refresh_token_expires_in) {
    throw new Error("eBay's token response had no refresh_token — cannot stay connected.");
  }

  const now = Date.now();
  await setStoredEbayToken({
    refreshToken: payload.refresh_token,
    refreshTokenExpiresAt: new Date(now + payload.refresh_token_expires_in * 1000).toISOString(),
    accessToken: payload.access_token,
    accessTokenExpiresAt: new Date(now + payload.expires_in * 1000).toISOString(),
    connectedAt: new Date(now).toISOString(),
  });
}

/**
 * Returns a currently-valid access token for calling eBay's APIs on
 * this shop's behalf, refreshing it first if the cached one is missing
 * or within 5 minutes of expiring (a small safety margin so a call
 * in flight doesn't expire mid-request). Returns null when this shop
 * has never connected its eBay account, or its refresh token has
 * itself expired (~18 months unused) — callers treat that the same as
 * "eBay isn't configured": skip the eBay side effect rather than fail
 * the whole request it's attached to (see components/AttachPartToJob.tsx).
 */
export async function getValidAccessToken(): Promise<string | null> {
  const stored = await getStoredEbayToken();
  if (!stored) {
    return null;
  }
  if (new Date(stored.refreshTokenExpiresAt).getTime() <= Date.now()) {
    return null;
  }

  const REFRESH_MARGIN_MS = 5 * 60 * 1000;
  const accessStillValid =
    stored.accessToken &&
    stored.accessTokenExpiresAt &&
    new Date(stored.accessTokenExpiresAt).getTime() - REFRESH_MARGIN_MS > Date.now();
  if (accessStillValid) {
    return stored.accessToken!;
  }

  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: stored.refreshToken,
  });
  const payload = await requestToken(body);
  const accessTokenExpiresAt = new Date(Date.now() + payload.expires_in * 1000).toISOString();
  await updateStoredAccessToken(payload.access_token, accessTokenExpiresAt);
  return payload.access_token;
}
