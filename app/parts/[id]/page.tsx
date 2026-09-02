import { ArrowLeft, ClipboardList, MapPin, Package } from "lucide-react";
import type { Metadata } from "next";
import Link from "next/link";
import { notFound } from "next/navigation";
import { JobStatusBadge, PartStatusBadge } from "@/components/StatusBadge";
import { getJobsForPart, getPartById } from "@/lib/inventory";
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

  const relatedJobs = getJobsForPart(part.id);

  return (
    <main className="space-y-6">
      <Link
        href="/"
        className="inline-flex items-center gap-1.5 text-sm font-semibold text-stone-600"
      >
        <ArrowLeft className="size-4" />
        All parts
      </Link>

      <section className="rounded-2xl border border-stone-200 bg-white p-5 shadow-sm">
        <div className="flex items-start justify-between gap-3">
          <div className="flex size-12 items-center justify-center rounded-xl bg-amber-50 text-amber-700">
            <Package className="size-6" />
          </div>
          <PartStatusBadge status={part.status} />
        </div>
        <p className="mt-4 font-mono text-xl font-semibold tracking-wide text-stone-900">
          {part.part_number}
        </p>
        <p className="mt-2 text-base leading-6 text-stone-600">{part.description}</p>
      </section>

      <section className="grid grid-cols-2 gap-3">
        <div className="rounded-2xl border border-stone-200 bg-white p-4">
          <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
            Bin location
          </p>
          <p className="mt-2 inline-flex items-center gap-1.5 text-lg font-semibold text-stone-900">
            <MapPin className="size-4 text-amber-700" />
            {part.bin_location}
          </p>
        </div>
        <div className="rounded-2xl border border-stone-200 bg-white p-4">
          <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
            On hand
          </p>
          <p className="mt-2 text-lg font-semibold text-stone-900">
            {part.quantity_on_hand}
          </p>
        </div>
      </section>

      <section className="space-y-3">
        <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">
          Jobs using this part
        </h2>
        {relatedJobs.length === 0 ? (
          <div className="rounded-2xl border border-dashed border-stone-300 bg-white px-4 py-8 text-center">
            <ClipboardList className="mx-auto size-7 text-stone-400" />
            <p className="mt-2 text-sm font-medium text-stone-700">
              Not assigned to an open job
            </p>
            <p className="mt-1 text-sm text-stone-500">
              This part is available for the next work order.
            </p>
          </div>
        ) : (
          <ul className="space-y-2">
            {relatedJobs.map((job) => (
              <li key={job.id}>
                <Link
                  href="/jobs"
                  className="block rounded-2xl border border-stone-200 bg-white p-4"
                >
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <p className="font-mono text-xs font-semibold text-stone-500">
                        Job {job.job_number}
                      </p>
                      <p className="mt-1 font-semibold text-stone-900">
                        {job.customer_name}
                      </p>
                      <p className="mt-1 text-sm text-stone-600">{job.appliance}</p>
                    </div>
                    <JobStatusBadge status={job.status} />
                  </div>
                </Link>
              </li>
            ))}
          </ul>
        )}
      </section>
    </main>
  );
}
