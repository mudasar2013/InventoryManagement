import type { Metadata } from "next";
import { notFound } from "next/navigation";
import { PartDetail } from "@/components/PartDetail";
import { getPartById } from "@/lib/inventory";
import { parts } from "@/lib/mockData";

type PartPageProps = PageProps<"/parts/[id]">;

export function generateStaticParams() {
  return parts.map((part) => ({ id: part.id }));
}

export async function generateMetadata({
  params,
}: PartPageProps): Promise<Metadata> {
  const { id } = await params;
  const part = getPartById(id);
  if (!part) {
    return { title: "Part not found" };
  }
  return { title: part.part_number };
}

export default async function PartDetailPage({ params }: PartPageProps) {
  const { id } = await params;
  const part = getPartById(id);

  if (!part) {
    notFound();
  }

  return <PartDetail partId={part.id} />;
}
