import { ChevronRight, MapPin, Package } from "lucide-react";
import Link from "next/link";
import type { Part } from "@/lib/types";
import { PartStatusBadge } from "./StatusBadge";

export function PartCard({ part }: { part: Part }) {
  return (
    <Link
      href={`/parts/${part.id}`}
      className="flex items-start gap-3 rounded-2xl border border-stone-200 bg-white p-4 shadow-sm transition-colors active:bg-stone-50"
    >
      <div className="flex size-11 shrink-0 items-center justify-center rounded-xl bg-amber-50 text-amber-700">
        <Package className="size-5" />
      </div>
      <div className="min-w-0 flex-1">
        <div className="flex items-start justify-between gap-2">
          <p className="font-mono text-sm font-semibold tracking-wide text-stone-900">
            {part.part_number}
          </p>
          <PartStatusBadge status={part.status} />
        </div>
        <p className="mt-1 text-sm leading-5 text-stone-600">{part.description}</p>
        <div className="mt-2 flex items-center gap-3 text-xs font-medium text-stone-500">
          <span className="inline-flex items-center gap-1">
            <MapPin className="size-3.5" />
            Bin {part.bin_location}
          </span>
          <span>{part.quantity_on_hand} on hand</span>
        </div>
      </div>
      <ChevronRight className="mt-3 size-4 shrink-0 text-stone-400" />
    </Link>
  );
}
