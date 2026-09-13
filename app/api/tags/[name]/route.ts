import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { deleteTag, isTagStoreConfigured, renameTag } from "@/lib/tags-store";

async function requireSession() {
  return getServerSession(authOptions);
}

/** Renames a tag everywhere it's used — in the vocabulary and on every
 *  part currently carrying it (see lib/tags-store.ts's renameTag). */
export async function PATCH(
  request: Request,
  { params }: { params: Promise<{ name: string }> },
) {
  const session = await requireSession();
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

  const { name } = await params;
  const body = await request.json().catch(() => null);
  const newName =
    body && typeof body === "object" ? (body as Record<string, unknown>).newName : null;
  if (typeof newName !== "string" || !newName.trim()) {
    return NextResponse.json({ error: "Missing or empty field: newName" }, { status: 400 });
  }

  try {
    const tags = await renameTag(decodeURIComponent(name), newName);
    return NextResponse.json({ tags });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Failed to rename tag." },
      { status: 500 },
    );
  }
}

export async function DELETE(
  _request: Request,
  { params }: { params: Promise<{ name: string }> },
) {
  const session = await requireSession();
  if (!session) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  const { name } = await params;
  try {
    const tags = await deleteTag(decodeURIComponent(name));
    return NextResponse.json({ tags });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Failed to delete tag." },
      { status: 500 },
    );
  }
}
