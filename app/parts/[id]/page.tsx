import type { Metadata } from "next";
import { getServerSession } from "next-auth/next";
import { notFound } from "next/navigation";
import { PartDetail } from "@/components/PartDetail";
import { authOptions } from "@/lib/auth/options";
import { loadInventory } from "@/lib/getInventory";
import { getPartById } from "@/lib/inventory";

type PartPageProps = PageProps<"/parts/[id]">;

// No generateStaticParams here on purpose: the catalog is read live from
// one or more inventory sources (see lib/getInventory.ts) and can change
// between deploys, so baking a fixed set of ids into the build would
// either 404 newly-added parts or ship stale ones. dynamicParams defaults
// to true, so every id renders on request against current data instead.

export async function generateMetadata({
  params,
}: PartPageProps): Promise<Metadata> {
  const { id } = await params;
  const session = await getServerSession(authOptions);
  const { parts } = await loadInventory(session?.accessToken);
  const part = getPartById(id, parts);
  if (!part) {
    return { title: "Part not found" };
  }
  return { title: part.part_number };
}

export default async function PartDetailPage({ params }: PartPageProps) {
  const { id } = await params;
  const session = await getServerSession(authOptions);
  const { parts } = await loadInventory(session?.accessToken);
  const part = getPartById(id, parts);

  if (!part) {
    notFound();
  }

  return <PartDetail partId={part.id} />;
}
