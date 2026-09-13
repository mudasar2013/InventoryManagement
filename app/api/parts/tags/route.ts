import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { isTagStoreConfigured, setTagsForPart } from "@/lib/tags-store";

/**
 * Sets a part's tags outright (see components/PartDetail.tsx's edit
 * form). Deliberately separate from PATCH /api/parts: tags are
 * app-native, not written to any SharePoint workbook, so this never
 * needs a Graph access token or a sourceId — only a signed-in session
 * and the part's own identity (part_number, its stable cross-source
 * key — see lib/tags-store.ts).
 */
export async function PATCH(request: Request) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  if (!isTagStoreConfigured()) {
    return NextResponse.json(
      {
        error:
          "No Redis store configured for this project. Add a Redis integration in Vercel (Storage tab) first.",
      },
      { status: 503 },
    );
  }

  const body = await request.json().catch(() => null);
  if (!body || typeof body !== "object") {
    return NextResponse.json({ error: "Invalid request body." }, { status: 400 });
  }
  const record = body as Record<string, unknown>;
  const partNumber = record.part_number;
  const tags = record.tags;
  if (typeof partNumber !== "string" || !partNumber.trim()) {
    return NextResponse.json({ error: "Missing or empty field: part_number" }, { status: 400 });
  }
  if (!Array.isArray(tags) || !tags.every((tag) => typeof tag === "string")) {
    return NextResponse.json({ error: "tags must be an array of strings." }, { status: 400 });
  }

  try {
    await setTagsForPart(partNumber.trim(), tags);
    return NextResponse.json({ ok: true });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Failed to update tags." },
      { status: 500 },
    );
  }
}
