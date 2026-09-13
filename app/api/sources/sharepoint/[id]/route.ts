import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { parseSharePointSourceInput } from "@/lib/sources/parse-sharepoint-input";
import {
  isSourceStoreConfigured,
  removeStoredSharePointSource,
  updateStoredSharePointSource,
} from "@/lib/sources/sharepoint-source-store";

export async function DELETE(
  _request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const { id } = await params;
  await removeStoredSharePointSource(id);
  return NextResponse.json({ ok: true });
}

/** Fixes a typo'd field on an existing source in place (same validation
 *  as adding one — see lib/sources/parse-sharepoint-input.ts). */
export async function PATCH(
  request: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  if (!isSourceStoreConfigured()) {
    return NextResponse.json(
      {
        error:
          "No Redis store configured for this project. Add a Redis integration in Vercel (Storage tab) first.",
      },
      { status: 503 },
    );
  }

  const { id } = await params;
  const body = await request.json().catch(() => null);
  const parsed = parseSharePointSourceInput(body);
  if ("error" in parsed) {
    return NextResponse.json({ error: parsed.error }, { status: 400 });
  }

  try {
    const source = await updateStoredSharePointSource(id, parsed.input);
    return NextResponse.json({ source });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Failed to update source." },
      { status: 500 },
    );
  }
}
