"use client";

import { ChevronRight, Database, Hash, MapPin, Package, Tag } from "lucide-react";
import Link from "next/link";
import { useInventory } from "@/components/InventoryProvider";
import { upnFromExtraFields, type Part } from "@/lib/types";
import { PartStatusBadge } from "./StatusBadge";

/** `disableLink` + `onClick` support the Parts page's bulk-select mode
 *  (see PartsExplorer.tsx): tapping a card there should toggle its
 *  checkbox, not navigate away to the part's own page. Renders a
 *  `<button>` instead of a `<Link>` in that mode — same visual card,
 *  different interaction. */
export function PartCard({
  part,
  disableLink = false,
  onClick,
}: {
  part: Part;
  disableLink?: boolean;
  onClick?: () => void;
}) {
  const { sourceLabels } = useInventory();
  const upn = upnFromExtraFields(part.extraFields);
  // A part reported by more than one source (see mergeParts) shows every
  // source it merged from, so "why does this row look different from the
  // sheet" is answerable at a glance instead of requiring a trip to
  // Settings → Data sources.
  const sourceLabel = (part.sourceIds ?? [])
    .map((id) => sourceLabels[id] ?? id)
    .join(" + ");
  const content = (
    <>
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
        {sourceLabel ? (
          <p className="mt-1 inline-flex items-center gap-1 text-[11px] font-medium text-sky-700">
            <Database className="size-3 shrink-0" />
            {sourceLabel}
          </p>
        ) : null}
        <dl className="mt-3 grid grid-cols-2 gap-2 text-xs">
          {upn ? (
            <div>
              <dt className="font-semibold uppercase tracking-wide text-stone-500">UPN#</dt>
              <dd className="mt-0.5 inline-flex items-center gap-1 font-medium text-stone-800">
                <Hash className="size-3.5 text-amber-700" />
                {upn}
              </dd>
            </div>
          ) : null}
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
        {part.tags && part.tags.length > 0 ? (
          <div className="mt-2 flex flex-wrap gap-1.5">
            {part.tags.map((tag) => (
              <span
                key={tag}
                className="inline-flex items-center gap-1 rounded-full bg-stone-100 px-2 py-0.5 text-[10px] font-semibold text-stone-600"
              >
                <Tag className="size-2.5" />
                {tag}
              </span>
            ))}
          </div>
        ) : null}
      </div>
    </>
  );

  if (disableLink) {
    return (
      <button
        type="button"
        onClick={onClick}
        className="flex w-full items-start gap-3 rounded-2xl border border-stone-200 bg-white p-4 text-left shadow-sm transition-colors active:bg-stone-50"
      >
        {content}
      </button>
    );
  }

  return (
    <Link
      href={`/parts/${part.id}`}
      className="flex items-start gap-3 rounded-2xl border border-stone-200 bg-white p-4 shadow-sm transition-colors active:bg-stone-50"
    >
      {content}
      <ChevronRight className="mt-3 size-4 shrink-0 text-stone-400" />
    </Link>
  );
}
