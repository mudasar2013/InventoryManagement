"use client";

import { CheckSquare, PackagePlus, PackageSearch, Search, Square, X } from "lucide-react";
import Link from "next/link";
import { useMemo, useState } from "react";
import { BulkEditPanel } from "@/components/BulkEditPanel";
import { LinkBanner } from "@/components/LinkBanner";
import { PartCard } from "@/components/PartCard";
import { countByStatus, searchParts } from "@/lib/inventory";
import type { PartStatus } from "@/lib/types";
import { useInventory } from "./InventoryProvider";

type Filter = "All" | PartStatus;

const filters: Filter[] = ["All", "In Stock", "Low Stock", "Out of Stock"];
const ALL = "All";

/** Distinct, non-blank values for one field across every part, sorted
 *  for a stable dropdown order. Shared by the category and location
 *  filters — both read whatever values already exist in the merged
 *  catalog rather than a separately curated list (see SettingsView). */
function distinctValues(values: (string | undefined)[]): string[] {
  const set = new Set(values.map((value) => value?.trim()).filter((value): value is string => Boolean(value)));
  return Array.from(set).sort((a, b) => a.localeCompare(b));
}

export function PartsExplorer() {
  const { parts, writableSources } = useInventory();
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<Filter>("All");
  const [categoryFilter, setCategoryFilter] = useState(ALL);
  const [locationFilter, setLocationFilter] = useState(ALL);
  const [tagFilter, setTagFilter] = useState(ALL);
  const [selecting, setSelecting] = useState(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkEditOpen, setBulkEditOpen] = useState(false);

  const counts = countByStatus(parts);
  const categories = useMemo(() => distinctValues(parts.map((part) => part.category)), [parts]);
  const locations = useMemo(() => distinctValues(parts.map((part) => part.bin_location)), [parts]);
  const tagValues = useMemo(() => distinctValues(parts.flatMap((part) => part.tags ?? [])), [parts]);

  const searched = searchParts(parts, query, filter);
  const visible = searched.filter((part) => {
    if (categoryFilter !== ALL && (part.category ?? "").trim() !== categoryFilter) return false;
    if (locationFilter !== ALL && part.bin_location.trim() !== locationFilter) return false;
    if (tagFilter !== ALL && !(part.tags ?? []).includes(tagFilter)) return false;
    return true;
  });

  const filterCount = (value: Filter) => {
    if (value === "All") return counts.all;
    if (value === "In Stock") return counts.inStock;
    if (value === "Low Stock") return counts.lowStock;
    return counts.outOfStock;
  };

  function toggleSelecting() {
    setSelecting((current) => !current);
    setSelectedIds(new Set());
  }

  function toggleSelected(id: string) {
    setSelectedIds((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  const selectedParts = parts.filter((part) => selectedIds.has(part.id));

  return (
    <div className="space-y-4 pb-16">
      <LinkBanner />

      <div className="flex items-center justify-between gap-2">
        {writableSources.length > 0 ? (
          <Link
            href="/parts/new"
            className="inline-flex items-center gap-1.5 text-sm font-semibold text-amber-700"
          >
            <PackagePlus className="size-4" />
            Add a part
          </Link>
        ) : (
          <span />
        )}
        <button
          type="button"
          onClick={toggleSelecting}
          className={`inline-flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 ${
            selecting
              ? "bg-stone-900 text-white ring-stone-900"
              : "bg-white text-stone-600 ring-stone-200"
          }`}
        >
          {selecting ? <X className="size-3.5" /> : <CheckSquare className="size-3.5" />}
          {selecting ? "Cancel" : "Select"}
        </button>
      </div>

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

      {categories.length > 0 || locations.length > 0 || tagValues.length > 0 ? (
        <div className="grid grid-cols-3 gap-2">
          {categories.length > 0 ? (
            <FilterSelect label="Category" value={categoryFilter} options={categories} onChange={setCategoryFilter} />
          ) : null}
          {locations.length > 0 ? (
            <FilterSelect label="Location" value={locationFilter} options={locations} onChange={setLocationFilter} />
          ) : null}
          {tagValues.length > 0 ? (
            <FilterSelect label="Tag" value={tagFilter} options={tagValues} onChange={setTagFilter} />
          ) : null}
        </div>
      ) : null}

      {visible.length === 0 ? (
        <div className="flex flex-col items-center rounded-2xl border border-dashed border-stone-300 bg-white px-6 py-12 text-center">
          <PackageSearch className="size-8 text-stone-400" />
          <h2 className="mt-3 text-base font-semibold text-stone-900">No parts match</h2>
          <p className="mt-1 text-sm text-stone-500">
            Try a different part number, or clear a filter.
          </p>
        </div>
      ) : (
        <ul className="space-y-3">
          {visible.map((part) => (
            <li key={part.id} className="flex items-center gap-2">
              {selecting ? (
                <button
                  type="button"
                  onClick={() => toggleSelected(part.id)}
                  aria-label={selectedIds.has(part.id) ? "Deselect part" : "Select part"}
                  className="shrink-0 text-stone-400"
                >
                  {selectedIds.has(part.id) ? (
                    <CheckSquare className="size-5 text-amber-700" />
                  ) : (
                    <Square className="size-5" />
                  )}
                </button>
              ) : null}
              <div className="min-w-0 flex-1">
                <PartCard part={part} disableLink={selecting} onClick={selecting ? () => toggleSelected(part.id) : undefined} />
              </div>
            </li>
          ))}
        </ul>
      )}

      {selecting && selectedIds.size > 0 ? (
        <div className="fixed inset-x-0 bottom-16 z-10 mx-auto flex max-w-lg items-center justify-between gap-3 border-t border-stone-200 bg-white/95 px-5 py-3 backdrop-blur-md sm:px-7">
          <span className="text-sm font-semibold text-stone-900">
            {selectedIds.size} selected
          </span>
          <button
            type="button"
            onClick={() => setBulkEditOpen(true)}
            className="rounded-xl bg-stone-900 px-4 py-2 text-sm font-semibold text-white"
          >
            Bulk edit
          </button>
        </div>
      ) : null}

      {bulkEditOpen ? (
        <BulkEditPanel
          parts={selectedParts}
          onClose={() => setBulkEditOpen(false)}
          onDone={() => {
            setBulkEditOpen(false);
            setSelecting(false);
            setSelectedIds(new Set());
          }}
        />
      ) : null}
    </div>
  );
}

function FilterSelect({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-[10px] font-semibold uppercase tracking-wide text-stone-500">
        {label}
      </span>
      <select
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className="h-9 w-full rounded-lg border border-stone-200 bg-white px-2 text-xs outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
      >
        <option value={ALL}>All</option>
        {options.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    </label>
  );
}
