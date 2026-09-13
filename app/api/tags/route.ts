import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { addTag, isTagStoreConfigured, listTags } from "@/lib/tags-store";

async function requireSession() {
  return getServerSession(authOptions);
}

export async function GET() {
  const session = await requireSession();
  if (!session) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }
  const tags = await listTags();
  return NextResponse.json({ tags });
}

/** Adds a tag to the vocabulary (see components/SettingsView.tsx). Tags
 *  are app-native — unlike /api/parts and /api/sources/sharepoint, this
 *  never touches Microsoft Graph, so no access token is required, only
 *  a signed-in session. */
export async function POST(request: Request) {
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

  const body = await request.json().catch(() => null);
  const name = body && typeof body === "object" ? (body as Record<string, unknown>).name : null;
  if (typeof name !== "string" || !name.trim()) {
    return NextResponse.json({ error: "Missing or empty field: name" }, { status: 400 });
  }

  try {
    const tags = await addTag(name);
    return NextResponse.json({ tags }, { status: 201 });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Failed to add tag." },
      { status: 500 },
    );
  }
}
