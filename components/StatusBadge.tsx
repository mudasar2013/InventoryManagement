import { AlertTriangle, CheckCircle2, XCircle } from "lucide-react";
import type { JobStatus, PartStatus } from "@/lib/types";

const partStyles: Record<PartStatus, string> = {
  "In Stock": "bg-emerald-50 text-emerald-800 ring-emerald-200",
  "Low Stock": "bg-amber-50 text-amber-800 ring-amber-200",
  "Out of Stock": "bg-rose-50 text-rose-800 ring-rose-200",
};

const jobStyles: Record<JobStatus, string> = {
  Scheduled: "bg-sky-50 text-sky-800 ring-sky-200",
  "In Progress": "bg-amber-50 text-amber-800 ring-amber-200",
  Completed: "bg-emerald-50 text-emerald-800 ring-emerald-200",
};

function PartIcon({ status }: { status: PartStatus }) {
  if (status === "In Stock") {
    return <CheckCircle2 className="size-3.5" aria-hidden />;
  }
  if (status === "Low Stock") {
    return <AlertTriangle className="size-3.5" aria-hidden />;
  }
  return <XCircle className="size-3.5" aria-hidden />;
}

export function PartStatusBadge({ status }: { status: PartStatus }) {
  return (
    <span
      className={`inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-semibold ring-1 ring-inset ${partStyles[status]}`}
    >
      <PartIcon status={status} />
      {status}
    </span>
  );
}

export function JobStatusBadge({ status }: { status: JobStatus }) {
  return (
    <span
      className={`inline-flex items-center rounded-full px-2.5 py-1 text-xs font-semibold ring-1 ring-inset ${jobStyles[status]}`}
    >
      {status}
    </span>
  );
}
