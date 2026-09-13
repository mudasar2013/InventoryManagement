import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { describeError, describeErrorDetail } from "@/lib/describe-error";
import { authOptions } from "@/lib/auth/options";
import { parsePartInput, parsePartUpdateInput } from "@/lib/sources/parse-part-input";
import { resolveSharePointConfigById } from "@/lib/sources/resolve-sharepoint-source";
import {
  addPartToWorkbook,
  COLUMN_MAP,
  connectToWorkbook,
  updatePartInWorkbook,
} from "@/lib/sources/sharepoint-excel-source";

async function requireAccessToken(): Promise<string | null> {
  const session = await getServerSession(authOptions);
  return session?.accessToken ?? null;
}

/**
 * Adds a brand-new part to a chosen SharePoint workbook (see
 * components/AddPartForm.tsx). Unlike editing, there's no existing row
 * to find — the fields go straight into a new row via
 * addPartToWorkbook.
 */
export async function POST(request: Request) {
  const accessToken = await requireAccessToken();
  if (!accessToken) {
    return NextResponse.json(
      { error: "Signed-in session has no Microsoft access token yet — sign out and back in." },
      { status: 401 },
    );
  }

  const body = await request.json().catch(() => null);
  const parsed = parsePartInput(body);
  if ("error" in parsed) {
    return NextResponse.json({ error: parsed.error }, { status: 400 });
  }

  const resolved = await resolveSharePointConfigById(parsed.input.sourceId);
  if (!resolved) {
    return NextResponse.json(
      { error: "That source no longer exists, or isn't a SharePoint workbook that can be written to." },
      { status: 404 },
    );
  }

  try {
    const { client, fileBase } = await connectToWorkbook(accessToken, resolved.config);
    await addPartToWorkbook(
      client,
      fileBase,
      resolved.config.tableName,
      resolved.config.columnMap ?? COLUMN_MAP,
      parsed.input.fields,
    );
    return NextResponse.json({ ok: true }, { status: 201 });
  } catch (error) {
    console.error("[api/parts] add failed:", error);
    return NextResponse.json(
      { error: describeError(error), debugDetail: describeErrorDetail(error) },
      { status: 502 },
    );
  }
}

/**
 * Updates an existing part's fields in its source workbook — finds the
 * row by `existingPartNumber` (the part number before this edit) and
 * overwrites the mapped columns with the submitted fields, which can
 * include a changed part_number (a rename). See
 * components/PartDetail.tsx for the calling UI.
 */
export async function PATCH(request: Request) {
  const accessToken = await requireAccessToken();
  if (!accessToken) {
    return NextResponse.json(
      { error: "Signed-in session has no Microsoft access token yet — sign out and back in." },
      { status: 401 },
    );
  }

  const body = await request.json().catch(() => null);
  const parsed = parsePartUpdateInput(body);
  if ("error" in parsed) {
    return NextResponse.json({ error: parsed.error }, { status: 400 });
  }

  const resolved = await resolveSharePointConfigById(parsed.input.sourceId);
  if (!resolved) {
    return NextResponse.json(
      { error: "That source no longer exists, or isn't a SharePoint workbook that can be written to." },
      { status: 404 },
    );
  }

  try {
    const { client, fileBase } = await connectToWorkbook(accessToken, resolved.config);
    await updatePartInWorkbook(
      client,
      fileBase,
      resolved.config.tableName,
      resolved.config.columnMap ?? COLUMN_MAP,
      parsed.input.existingPartNumber,
      parsed.input.fields,
      parsed.input.existingPartOccurrence,
    );
    return NextResponse.json({ ok: true });
  } catch (error) {
    console.error("[api/parts] update failed:", error);
    return NextResponse.json(
      { error: describeError(error), debugDetail: describeErrorDetail(error) },
      { status: 502 },
    );
  }
}
