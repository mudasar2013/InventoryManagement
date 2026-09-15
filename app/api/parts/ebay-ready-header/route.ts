import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { describeError, describeErrorDetail } from "@/lib/describe-error";
import { resolveSharePointConfigById } from "@/lib/sources/resolve-sharepoint-source";
import { connectToWorkbook, peekEbayReadyHeader } from "@/lib/sources/sharepoint-excel-source";

/**
 * Previews the real "Ebay ready" column header `sourceId`'s sheet uses,
 * if it has one — see peekEbayReadyHeader. Used by
 * components/AddPartForm.tsx so the Ebay-ready dropdown writes to
 * whatever this source actually calls that column ("Ebay Ready
 * (Yes/No)" on most sources, "Ebay Ready (Yes/N0)" — a typo — on one)
 * instead of a single hardcoded literal that silently matched nothing
 * on sources spelled differently.
 *
 * Returns `{ header: null }` — not an error — whenever there's simply
 * nothing to preview: no Ebay-ready-shaped column on this source's
 * sheet, or the source can't be resolved at all. The form treats this
 * the same way it treats a missing UPN# column: don't show the field.
 */
export async function GET(request: Request) {
  const session = await getServerSession(authOptions);
  const accessToken = session?.accessToken;
  if (!accessToken) {
    return NextResponse.json(
      { error: "Signed-in session has no Microsoft access token yet — sign out and back in." },
      { status: 401 },
    );
  }

  const sourceId = new URL(request.url).searchParams.get("sourceId");
  if (!sourceId) {
    return NextResponse.json({ error: "sourceId is required." }, { status: 400 });
  }

  const resolved = await resolveSharePointConfigById(sourceId);
  if (!resolved) {
    return NextResponse.json({ header: null });
  }

  try {
    const { client, fileBase } = await connectToWorkbook(accessToken, resolved.config);
    const header = await peekEbayReadyHeader(client, fileBase, resolved.config.tableName);
    return NextResponse.json({ header: header ?? null });
  } catch (error) {
    console.error("[api/parts/ebay-ready-header] preview failed:", error);
    return NextResponse.json(
      { error: describeError(error), debugDetail: describeErrorDetail(error) },
      { status: 502 },
    );
  }
}
