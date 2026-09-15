import { createHash } from "crypto";
import { NextResponse } from "next/server";

/**
 * eBay's required "Marketplace Account Deletion/Closure Notifications"
 * endpoint. eBay requires every app using its APIs to implement this —
 * and keeps the developer keyset disabled until it's configured and
 * verified — even for an app like this one that never stores any eBay
 * buyer/seller personal data. See
 * https://developer.ebay.com/develop/guides/sell/marketplace-user-account-deletion.
 *
 * GET handles eBay's one-time (and ongoing, on re-verify) endpoint
 * challenge: hash challengeCode + verificationToken + this endpoint's
 * own URL, in that exact order, with SHA-256, and echo the hex digest
 * back as {challengeResponse}. ENDPOINT_URL below must be byte-for-byte
 * identical to whatever URL is entered in the eBay Developer Portal's
 * notification settings — a trailing slash or http/https mismatch
 * fails verification even though the endpoint itself is reachable.
 *
 * POST delivers the actual notification once a user asks eBay to
 * delete their data. This app has never stored any eBay user's
 * personal data — it only ever calls eBay's own listing APIs for this
 * shop's own inventory, never buyer/seller PII — so there's nothing
 * here to act on. Acknowledging with a 2xx within eBay's 24-hour
 * window is the entire obligation; the payload is logged for a paper
 * trail in case eBay's compliance team ever asks.
 */
const ENDPOINT_URL = "https://inventory.homeappliancecare.us/api/ebay/account-deletion";

export async function GET(request: Request) {
  const challengeCode = new URL(request.url).searchParams.get("challenge_code");
  const verificationToken = process.env.EBAY_VERIFICATION_TOKEN;

  if (!challengeCode) {
    return NextResponse.json({ error: "challenge_code is required." }, { status: 400 });
  }
  if (!verificationToken) {
    return NextResponse.json(
      { error: "EBAY_VERIFICATION_TOKEN is not configured on this deployment." },
      { status: 500 },
    );
  }

  const hash = createHash("sha256");
  hash.update(challengeCode);
  hash.update(verificationToken);
  hash.update(ENDPOINT_URL);
  const challengeResponse = hash.digest("hex");

  return NextResponse.json({ challengeResponse });
}

export async function POST(request: Request) {
  try {
    const payload = await request.text();
    console.log("[api/ebay/account-deletion] notification received:", payload);
  } catch (error) {
    console.error("[api/ebay/account-deletion] failed to read notification body:", error);
  }
  // Nothing to actually delete — this app never stores eBay user PII —
  // acknowledging is the whole job. 204 is one of eBay's accepted codes.
  return new NextResponse(null, { status: 204 });
}
