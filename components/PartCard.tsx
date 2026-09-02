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
          <div>
            <p className="text-[11px] font-semibold uppercase tracking-wide text-stone-500">
              Part number
            </p>
            <p className="font-mono text-sm font-semibold tracking-wide text-stone-900">
              {part.part_number}
            </p>
          </div>
          <PartStatusBadge status={part.status} />
        </div>
        <p className="mt-1 text-sm leading-5 text-stone-600">{part.description}</p>
        <dl className="mt-3 grid grid-cols-2 gap-2 text-xs">
          <div>
            <dt className="font-semibold uppercase tracking-wide text-stone-500">
              Bin location
            </dt>
            <dd className="mt-0.5 inline-flex items-center gap-1 font-medium text-stone-800">
              <MapPin className="size-3.5 text-amber-700" />
              {part.bin_location}
            </dd>
          </div>
          <div>
            <dt className="font-semibold uppercase tracking-wide text-stone-500">
              Quantity on hand
            </dt>
            <dd className="mt-0.5 font-medium text-stone-800">{part.quantity_on_hand}</dd>
          </div>
        </dl>
      </div>
      <ChevronRight className="mt-3 size-4 shrink-0 text-stone-400" />
    </Link>
  );
}
