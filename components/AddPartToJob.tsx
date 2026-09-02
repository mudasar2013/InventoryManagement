"use client";

import { Plus, Search, X } from "lucide-react";
import { useState } from "react";
import { PartStatusBadge } from "@/components/StatusBadge";
import { searchParts } from "@/lib/inventory";
import type { Job } from "@/lib/types";
import { useInventory } from "./InventoryProvider";

export function AddPartToJob({ job }: { job: Job }) {
  const { parts, isLinked, linkPart } = useInventory();
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");

  const matches = !open
    ? []
    : searchParts(parts, query).filter((part) => !isLinked(job.id, part.id));

  if (!open) {
    return (
      <button
        type="button"
        onClick={() => setOpen(true)}
        className="mt-3 inline-flex w-full items-center justify-center gap-1.5 rounded-xl border border-dashed border-stone-300 px-3 py-2.5 text-sm font-semibold text-stone-700"
      >
        <Plus className="size-4" />
        Add part by number
      </button>
    );
  }

  return (
    <div className="mt-3 rounded-xl border border-stone-200 bg-stone-50 p-3">
      <div className="flex items-center justify-between gap-2">
        <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
          Search part number
        </p>
        <button
          type="button"
          onClick={() => {
            setOpen(false);
            setQuery("");
          }}
          className="rounded-full p-1 text-stone-400"
          aria-label="Close part search"
        >
          <X className="size-4" />
        </button>
      </div>
      <label className="relative mt-2 block">
        <span className="sr-only">Search by part number</span>
        <Search className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-stone-400" />
        <input
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Search by part number"
          autoFocus
          className="h-11 w-full rounded-xl border border-stone-200 bg-white pl-9 pr-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
        />
      </label>
      <ul className="mt-2 space-y-2">
        {matches.length === 0 ? (
          <li className="px-1 py-3 text-center text-sm text-stone-500">
            {query
              ? "No unmatched parts for that number."
              : "Every catalog part is already on this job, or start typing a number."}
          </li>
        ) : (
          matches.map((part) => (
            <li key={part.id}>
              <button
                type="button"
                onClick={() => {
                  linkPart(job.id, part.id);
                  setOpen(false);
                  setQuery("");
                }}
                className="flex w-full items-start justify-between gap-3 rounded-xl bg-white px-3 py-2.5 text-left"
              >
                <span>
                  <span className="block font-mono text-xs font-semibold text-stone-900">
                    {part.part_number}
                  </span>
                  <span className="mt-0.5 block text-xs text-stone-500">
                    Bin {part.bin_location} · {part.quantity_on_hand} on hand
                  </span>
                </span>
                <PartStatusBadge status={part.status} />
              </button>
            </li>
          ))
        )}
      </ul>
    </div>
  );
}
