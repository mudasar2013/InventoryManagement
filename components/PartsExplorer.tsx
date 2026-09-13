"use client";

import { PackagePlus, PackageSearch, Search, X } from "lucide-react";
import Link from "next/link";
import { useState } from "react";
import { LinkBanner } from "@/components/LinkBanner";
import { PartCard } from "@/components/PartCard";
import { countByStatus, searchParts } from "@/lib/inventory";
import type { PartStatus } from "@/lib/types";
import { useInventory } from "./InventoryProvider";

type Filter = "All" | PartStatus;

const filters: Filter[] = ["All", "In Stock", "Low Stock", "Out of Stock"];

export function PartsExplorer() {
  const { parts, writableSources } = useInventory();
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<Filter>("All");
  const counts = countByStatus(parts);

  const visible = searchParts(parts, query, filter);

  const filterCount = (value: Filter) => {
    if (value === "All") return counts.all;
    if (value === "In Stock") return counts.inStock;
    if (value === "Low Stock") return counts.lowStock;
    return counts.outOfStock;
  };

  return (
    <div className="space-y-4">
      <LinkBanner />

      {writableSources.length > 0 ? (
        <Link
          href="/parts/new"
          className="inline-flex items-center gap-1.5 text-sm font-semibold text-amber-700"
        >
          <PackagePlus className="size-4" />
          Add a part
        </Link>
      ) : null}

      <label className="relative block">
        <span className="sr-only">Search by part number</span>
        <Search className="pointer-events-none absolute left-3.5 top-1/2 size-4 -translate-y-1/2 text-stone-400" />
        <input
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Search by part number"
          inputMode="search"
          autoComplete="off"
          autoCapitalize="characters"
          className="h-12 w-full rounded-2xl border border-stone-200 bg-white pl-10 pr-11 text-sm text-stone-900 shadow-sm outline-none placeholder:text-stone-400 focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
        />
        {query ? (
          <button
            type="button"
            onClick={() => setQuery("")}
            className="absolute right-3 top-1/2 -translate-y-1/2 rounded-full p-1 text-stone-400 hover:bg-stone-100 hover:text-stone-600"
            aria-label="Clear search"
          >
            <X className="size-4" />
          </button>
        ) : null}
      </label>
      <p className="text-xs text-stone-500">
        Look up stock from the shelf — no job required. Description and bin also match.
      </p>

      <div className="-mx-1 flex gap-2 overflow-x-auto px-1 pb-1">
        {filters.map((value) => {
          const selected = filter === value;
          return (
            <button
              key={value}
              type="button"
              onClick={() => setFilter(value)}
              className={`shrink-0 rounded-full px-3.5 py-2 text-xs font-semibold ring-1 ${
                selected
                  ? "bg-stone-900 text-white ring-stone-900"
                  : "bg-white text-stone-600 ring-stone-200"
              }`}
            >
              {value} ({filterCount(value)})
            </button>
          );
        })}
      </div>

      {visible.length === 0 ? (
        <div className="flex flex-col items-center rounded-2xl border border-dashed border-stone-300 bg-white px-6 py-12 text-center">
          <PackageSearch className="size-8 text-stone-400" />
          <h2 className="mt-3 text-base font-semibold text-stone-900">No parts match</h2>
          <p className="mt-1 text-sm text-stone-500">
            Try a different part number, or clear the stock filter.
          </p>
        </div>
      ) : (
        <ul className="space-y-3">
          {visible.map((part) => (
            <li key={part.id}>
              <PartCard part={part} />
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
