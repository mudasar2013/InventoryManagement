import { getServerSession } from "next-auth/next";
import { NextResponse } from "next/server";
import { authOptions } from "@/lib/auth/options";

/**
 * The "Suggested" half of the eBay description panel (see
 * components/PartDetail.tsx and app/api/ebay/description, the
 * "Actual" half). Drafts an eBay-ready listing description from a
 * part's own catalog fields via the Anthropic API — needs
 * ANTHROPIC_API_KEY set as a server env var (a Vercel Production env
 * var, same as the EBAY_* ones); ANTHROPIC_MODEL optionally overrides
 * the model used. Returns a clear error rather than throwing when
 * that key isn't configured, so the button in the UI can say exactly
 * what's missing instead of a generic failure.
 */

interface SuggestDescriptionInput {
  part_number: string;
  description?: string;
  category?: string;
  condition?: string;
}

function parseInput(body: unknown): { input: SuggestDescriptionInput } | { error: string } {
  if (!body || typeof body !== "object") {
    return { error: "Invalid request body." };
  }
  const record = body as Record<string, unknown>;
  if (typeof record.part_number !== "string" || !record.part_number.trim()) {
    return { error: "Missing or empty field: part_number" };
  }
  return {
    input: {
      part_number: record.part_number.trim(),
      description: typeof record.description === "string" ? record.description.trim() : undefined,
      category: typeof record.category === "string" ? record.category.trim() : undefined,
      condition: typeof record.condition === "string" ? record.condition.trim() : undefined,
    },
  };
}

function buildPrompt(input: SuggestDescriptionInput): string {
  const lines = [
    `Part number (SKU): ${input.part_number}`,
    input.description ? `Internal description: ${input.description}` : null,
    input.category ? `Category: ${input.category}` : null,
    input.condition ? `Condition: ${input.condition}` : null,
  ].filter((line): line is string => Boolean(line));

  return [
    "Write an eBay listing description for a genuine appliance repair part sold by a small appliance-repair business.",
    "Use the catalog details below — don't invent specs, brands, or compatibility claims that aren't given.",
    "Keep it factual, scannable (short paragraphs and/or a short bullet list of key points), and written to help a buyer confirm this is the right part for their appliance.",
    "Do not include pricing, shipping, or return-policy text — those are handled elsewhere on the listing.",
    "",
    ...lines,
  ].join("\n");
}

export async function POST(request: Request) {
  const session = await getServerSession(authOptions);
  if (!session) {
    return NextResponse.json({ error: "Sign in first." }, { status: 401 });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return NextResponse.json(
      { error: "ANTHROPIC_API_KEY is not configured on the server." },
      { status: 501 },
    );
  }

  const body = await request.json().catch(() => null);
  const parsed = parseInput(body);
  if ("error" in parsed) {
    return NextResponse.json({ error: parsed.error }, { status: 400 });
  }

  const model = process.env.ANTHROPIC_MODEL || "claude-sonnet-4-5";
  let response: Response;
  try {
    response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model,
        max_tokens: 700,
        messages: [{ role: "user", content: buildPrompt(parsed.input) }],
      }),
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Anthropic API request failed." },
      { status: 502 },
    );
  }

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const message = payload?.error?.message || `Anthropic API returned ${response.status}`;
    return NextResponse.json({ error: message }, { status: 502 });
  }

  const text = Array.isArray(payload?.content)
    ? payload.content
        .filter((block: unknown) => (block as Record<string, unknown>)?.type === "text")
        .map((block: unknown) => (block as Record<string, unknown>).text as string)
        .join("\n")
    : "";

  if (!text.trim()) {
    return NextResponse.json({ error: "Anthropic API returned an empty response." }, { status: 502 });
  }

  return NextResponse.json({ suggested: text.trim() });
}
