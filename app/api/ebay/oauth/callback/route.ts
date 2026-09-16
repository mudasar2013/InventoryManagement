import { NextResponse } from "next/server";
import { describeError } from "@/lib/describe-error";
import { exchangeCodeForTokens } from "@/lib/ebay/oauth";

/**
 * Where eBay redirects the shop owner back to after they sign in and
 * grant this app consent (registered as this app's "auth accepted
 * URL" under the RuName's sign-in settings — see lib/ebay/config.ts's
 * EBAY_RU_NAME doc comment). Trades the one-time `code` for a real
 * access + refresh token pair (see exchangeCodeForTokens) and shows a
 * plain confirmation page — this route isn't part of the signed-in app
 * shell (see app/layout.tsx), since eBay's redirect has no session
 * cookie of its own, so a small standalone HTML page is simpler than
 * trying to render through the app's normal pages.
 *
 * eBay also redirects here (same registered URL) if the shop owner
 * declines consent — the "declined" URL falls back to this same page
 * since it's simplest to just say so, rather than maintaining a
 * second near-identical route for the decline case.
 */
function htmlPage(title: string, message: string): NextResponse {
  return new NextResponse(
    `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title>
    <style>body{font-family:system-ui,sans-serif;background:#1c1917;color:#e7e5e4;display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0;padding:24px}
    main{max-width:420px;text-align:center}
    h1{font-size:1.25rem;margin-bottom:.5rem}
    p{color:#a8a29e;line-height:1.5}
    a{color:#fbbf24}</style></head>
    <body><main><h1>${title}</h1><p>${message}</p><p><a href="/">Back to the app</a></p></main></body></html>`,
    { status: 200, headers: { "Content-Type": "text/html" } },
  );
}

export async function GET(request: Request) {
  const url = new URL(request.url);
  const code = url.searchParams.get("code");
  const error = url.searchParams.get("error") ?? url.searchParams.get("isAuthSuccessful");

  if (!code) {
    return htmlPage(
      "eBay sign-in declined",
      error
        ? `eBay reported: ${error}. Nothing was connected — try again from the Developer Portal's "Sign in to Production for OAuth" button.`
        : "No authorization code came back from eBay, so nothing was connected.",
    );
  }

  try {
    await exchangeCodeForTokens(code);
  } catch (err) {
    console.error("[api/ebay/oauth/callback] token exchange failed:", err);
    return htmlPage(
      "eBay connection failed",
      `The sign-in worked, but saving the token failed: ${describeError(err)}. Nothing is connected yet — this can be retried.`,
    );
  }

  return htmlPage(
    "eBay connected",
    "This shop's eBay account is now connected. You can close this tab and go back to the app.",
  );
}
