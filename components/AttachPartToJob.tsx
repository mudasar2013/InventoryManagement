"use client";

import { Link2, Plus } from "lucide-react";
import { JobStatusBadge } from "@/components/StatusBadge";
import type { Part } from "@/lib/types";
import { useInventory } from "./InventoryProvider";

export function AttachPartToJob({ part }: { part: Part }) {
  const { jobs, isLinked, linkPart } = useInventory();

  return (
    <section className="space-y-3">
      <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">
        Attach to a job
      </h2>
      <p className="text-sm text-stone-600">
        Linking reserves the part on the work order. On-hand quantity stays at{" "}
        <span className="font-semibold text-stone-900">{part.quantity_on_hand}</span>.
      </p>
      <ul className="space-y-2">
        {jobs.map((job) => {
          const linked = isLinked(job.id, part.id);
          return (
            <li
              key={job.id}
              className="flex items-center justify-between gap-3 rounded-2xl border border-stone-200 bg-white p-3"
            >
              <div className="min-w-0">
                <p className="font-mono text-xs font-semibold text-stone-500">
                  Job {job.job_number}
                </p>
                <p className="truncate text-sm font-semibold text-stone-900">
                  {job.customer_name}
                </p>
                <div className="mt-1">
                  <JobStatusBadge status={job.status} />
                </div>
              </div>
              <button
                type="button"
                onClick={() => linkPart(job.id, part.id)}
                disabled={linked}
                className={`inline-flex shrink-0 items-center gap-1.5 rounded-full px-3 py-2 text-xs font-semibold ${
                  linked
                    ? "bg-stone-100 text-stone-500"
                    : "bg-stone-900 text-white"
                }`}
              >
                {linked ? <Link2 className="size-3.5" /> : <Plus className="size-3.5" />}
                {linked ? "On job" : "Attach"}
              </button>
            </li>
          );
        })}
      </ul>
    </section>
  );
}
