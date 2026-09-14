import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { describeError, describeErrorDetail } from "@/lib/describe-error";
import { resolveSharePointConfigById } from "@/lib/sources/resolve-sharepoint-source";
import { connectToWorkbook, peekNextUpn } from "@/lib/sources/sharepoint-excel-source";

/**
 * Previews the UPN# a new part added to `sourceId` would be auto-numbered
 * with — see peekNextUpn. Used by components/AddPartForm.tsx to show the
 * next free UPN# as soon as the form opens (and again whenever the
 * technician switches sources), rather than only after saving.
 *
 * Returns `{ header: null }` — not an error — whenever there's simply
 * nothing to preview: no UPN-shaped column on this source's sheet, its
 * last entry isn't a plain number, or the source can't be resolved at
 * all. The form treats all of these the same way: don't show a UPN
 * field, exactly as if this route didn't exist.
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
    const preview = await peekNextUpn(client, fileBase, resolved.config.tableName);
    return NextResponse.json(preview ? { header: preview.header, next: preview.next } : { header: null });
  } catch (error) {
    console.error("[api/parts/next-upn] preview failed:", error);
    return NextResponse.json(
      { error: describeError(error), debugDetail: describeErrorDetail(error) },
      { status: 502 },
    );
  }
}
