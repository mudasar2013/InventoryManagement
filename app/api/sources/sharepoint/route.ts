import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";
import { parseSharePointSourceInput } from "@/lib/sources/parse-sharepoint-input";
import {
  addStoredSharePointSource,
  isSourceStoreConfigured,
  listStoredSharePointSources,
} from "@/lib/sources/sharepoint-source-store";

// Every route in this app is already gated by proxy.ts (the auth
// middleware matches everything except /api/auth), so an unauthenticated
// request never reaches here — this second check is only because a
// missing session (e.g. an expired token proxy.ts still let through)
// should fail loudly on a write endpoint rather than silently act as
// nobody.
async function requireSession() {
  const session = await getServerSession(authOptions);
  if (!session) {
    return null;
  }
  return session;
}

export async function GET() {
  const session = await requireSession();
  if (!session) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const sources = await listStoredSharePointSources();
  return NextResponse.json({ sources });
}

export async function POST(request: Request) {
  const session = await requireSession();
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

  const body = await request.json().catch(() => null);
  const parsed = parseSharePointSourceInput(body);
  if ("error" in parsed) {
    return NextResponse.json({ error: parsed.error }, { status: 400 });
  }

  try {
    const source = await addStoredSharePointSource(parsed.input);
    return NextResponse.json({ source }, { status: 201 });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Failed to add source." },
      { status: 500 },
    );
  }
}
