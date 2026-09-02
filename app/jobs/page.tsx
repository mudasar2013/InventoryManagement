import { AlertTriangle, CalendarClock, MapPin, Package, Wrench } from "lucide-react";
import Link from "next/link";
import { AppHeader } from "@/components/AppHeader";
import { JobStatusBadge, PartStatusBadge } from "@/components/StatusBadge";
import { getJobPartsForJob, getPartById } from "@/lib/inventory";
import { jobs } from "@/lib/mockData";

function formatWhen(value: string) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "short",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(value));
}

export default function JobsPage() {
  return (
    <main className="space-y-6">
      <AppHeader
        title="Service jobs"
        subtitle="Parts already staged for today's calls. Out-of-stock items need a pickup before you roll."
      />

      {jobs.length === 0 ? (
        <div className="rounded-2xl border border-dashed border-stone-300 bg-white px-6 py-12 text-center">
          <Wrench className="mx-auto size-8 text-stone-400" />
          <h2 className="mt-3 text-base font-semibold text-stone-900">No open jobs</h2>
          <p className="mt-1 text-sm text-stone-500">
            New work orders will show the parts they need here.
          </p>
        </div>
      ) : (
        <ul className="space-y-4">
          {jobs.map((job) => {
            const required = getJobPartsForJob(job.id);
            const missing = required.some((item) => {
              const part = getPartById(item.part_id);
              return !part || part.quantity_on_hand < item.quantity_needed;
            });

            return (
              <li
                key={job.id}
                className="rounded-2xl border border-stone-200 bg-white p-4 shadow-sm"
              >
                <div className="flex items-start justify-between gap-3">
                  <div>
                    <p className="font-mono text-xs font-semibold uppercase tracking-wide text-stone-500">
                      Job {job.job_number}
                    </p>
                    <h2 className="mt-1 text-lg font-semibold text-stone-900">
                      {job.customer_name}
                    </h2>
                    <p className="mt-1 text-sm text-stone-600">{job.appliance}</p>
                  </div>
                  <JobStatusBadge status={job.status} />
                </div>

                <p className="mt-3 text-sm leading-6 text-stone-700">{job.issue}</p>

                <p className="mt-3 inline-flex items-center gap-1.5 text-xs font-medium text-stone-500">
                  <CalendarClock className="size-3.5" />
                  {formatWhen(job.scheduled_for)}
                </p>

                {missing ? (
                  <p className="mt-3 inline-flex items-center gap-1.5 rounded-full bg-rose-50 px-2.5 py-1 text-xs font-semibold text-rose-800">
                    <AlertTriangle className="size-3.5" />
                    Short parts
                  </p>
                ) : null}

                <ul className="mt-4 space-y-2">
                  {required.map((item) => {
                    const part = getPartById(item.part_id);
                    if (!part) {
                      return (
                        <li
                          key={item.id}
                          className="rounded-xl bg-stone-50 px-3 py-2 text-sm text-stone-500"
                        >
                          Unknown part
                        </li>
                      );
                    }

                    return (
                      <li key={item.id}>
                        <Link
                          href={`/parts/${part.id}`}
                          className="flex items-center justify-between gap-3 rounded-xl bg-stone-50 px-3 py-2.5"
                        >
                          <span className="min-w-0">
                            <span className="flex items-center gap-1.5 font-mono text-xs font-semibold text-stone-900">
                              <Package className="size-3.5 text-amber-700" />
                              {part.part_number}
                            </span>
                            <span className="mt-0.5 block truncate text-xs text-stone-500">
                              Need {item.quantity_needed} · {part.description}
                            </span>
                          </span>
                          <span className="flex shrink-0 flex-col items-end gap-1">
                            <PartStatusBadge status={part.status} />
                            <span className="inline-flex items-center gap-1 text-[11px] font-medium text-stone-500">
                              <MapPin className="size-3" />
                              {part.bin_location}
                            </span>
                          </span>
                        </Link>
                      </li>
                    );
                  })}
                </ul>
              </li>
            );
          })}
        </ul>
      )}
    </main>
  );
}
