/**
 * eBay application credentials and environment (Sandbox vs Production)
 * — see the eBay Developer Program's "Application Keys" page. Kept in
 * one place so lib/ebay/oauth.ts and the future listing-management
 * calls (decrement/end a listing when a part is used — see
 * components/AttachPartToJob.tsx) agree on which environment and
 * credentials they're talking to.
 *
 * Required environment variables:
 *   EBAY_APP_ID    — "App ID (Client ID)" from the Developer Portal
 *   EBAY_CERT_ID   — "Cert ID (Client Secret)" — keep this secret
 *   EBAY_DEV_ID    — "Dev ID"
 *   EBAY_RU_NAME   — the redirect/RuName created under "User Tokens
 *                    (eBay Sign-In)" → "Get a Token from eBay via Your
 *                    Application". This is a short eBay-assigned name
 *                    (e.g. "Mudssar_Syed-MudssarS-Invent-jqnfb"), NOT
 *                    the literal https:// callback URL — eBay's token
 *                    exchange wants the RuName in `redirect_uri`.
 *   EBAY_ENV       — "production" or "sandbox". Defaults to
 *                    "production" since that's the only keyset this
 *                    shop has created so far (see the Application Keys
 *                    page: no Sandbox keyset exists yet). Set this to
 *                    "sandbox" once a Sandbox keyset exists, to test
 *                    against it without touching real listings.
 */

export type EbayEnvironment = "production" | "sandbox";

export interface EbayConfig {
  appId: string;
  certId: string;
  devId: string;
  ruName: string;
  env: EbayEnvironment;
}

export function readEbayConfig(): EbayConfig | null {
  const appId = process.env.EBAY_APP_ID;
  const certId = process.env.EBAY_CERT_ID;
  const devId = process.env.EBAY_DEV_ID;
  const ruName = process.env.EBAY_RU_NAME;
  if (!appId || !certId || !devId || !ruName) {
    return null;
  }
  const env: EbayEnvironment = process.env.EBAY_ENV === "sandbox" ? "sandbox" : "production";
  return { appId, certId, devId, ruName, env };
}

/** eBay's OAuth token endpoint — issues and refreshes user access
 *  tokens. Same host pattern for both environments, "sandbox." only
 *  inserted for the Sandbox case. */
export function ebayOAuthTokenUrl(env: EbayEnvironment): string {
  return env === "sandbox"
    ? "https://api.sandbox.ebay.com/identity/v1/oauth2/token"
    : "https://api.ebay.com/identity/v1/oauth2/token";
}

/** eBay's OAuth user-consent page — where a seller is sent to sign in
 *  and grant this app access. Not used by the callback route itself,
 *  but kept alongside the token URL for whatever initiates that
 *  redirect later (today, that's the "Sign in to Production for
 *  OAuth" button on eBay's own Developer Portal — see
 *  app/ebay/page.tsx's doc comment once that page exists). */
export function ebayOAuthAuthorizeUrl(env: EbayEnvironment): string {
  return env === "sandbox"
    ? "https://auth.sandbox.ebay.com/oauth2/authorize"
    : "https://auth.ebay.com/oauth2/authorize";
}

/** eBay's Trading API endpoint (XML) — used for SKU-based inventory
 *  revision and ending listings, see lib/ebay/trading.ts. */
export function ebayTradingApiUrl(env: EbayEnvironment): string {
  return env === "sandbox"
    ? "https://api.sandbox.ebay.com/ws/api.dll"
    : "https://api.ebay.com/ws/api.dll";
}
