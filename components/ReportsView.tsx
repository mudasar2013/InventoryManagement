"use client";

import { Download, Printer } from "lucide-react";
import { useMemo, useState } from "react";
import { AppHeader } from "@/components/AppHeader";
import { useInventory } from "@/components/InventoryProvider";
import { distinctValues } from "@/lib/inventory";
import {
  discoverExtraColumns,
  extraColumnValues,
  matchesNumericFilter,
  matchesOptionFilter,
  matchesTextFilter,
  toCsv,
  type NumericRange,
} from "@/lib/reports";
import { upnFromExtraFields, type Part } from "@/lib/types";

/**
 * A generic filter-by-any-column report over the merged parts catalog
 * (see the user's request: "find out different filtered lists on all
 * column headers and their combinations"). Every column narrows the
 * same result set — filters combine with AND only, never OR — so
 * "Category = Motors AND Status = Low Stock" is answerable, but not
 * "Category = Motors OR Category = Pumps" in one report.
 *
 * Column model: `standardColumns` are the fields every part has
 * regardless of source (plus Tags and UPN#, which are app-managed/
 * commonly-present rather than universal); `dynamicColumns` are
 * whatever extra sheet headers actually show up in the catalog right
 * now (Condition, "Ebay Ready (Yes/No)", "Entry Date", "UPC#", ...) —
 * see discoverExtraColumns. The 7 always-on standard columns
 * (part_number, description, category, bin_location, quantity_on_hand,
 * status, source) are always shown in the results table; Tags, UPN#,
 * and every sheet column are only added to the table once the user has
 * an active filter on them, so the table doesn't start impossibly wide.
 */

type ReportColumn =
  | { key: string; label: string; kind: "text"; alwaysVisible: boolean; get: (part: Part) => string }
  | { key: string; label: string; kind: "options"; alwaysVisible: boolean; get: (part: Part) => string[] }
  | { key: string; label: string; kind: "numeric"; alwaysVisible: boolean; get: (part: Part) => number };

type NumericFilterInput = { min: string; max: string };

function displayValue(column: ReportColumn, part: Part): string {
  if (column.kind === "text") return column.get(part);
  if (column.kind === "options") return column.get(part).join(", ");
  return String(column.get(part));
}

function parseRange(raw: NumericFilterInput | undefined): NumericRange {
  return {
    min: raw?.min.trim() ? Number(raw.min) : undefined,
    max: raw?.max.trim() ? Number(raw.max) : undefined,
  };
}

function matchesColumn(
  column: ReportColumn,
  part: Part,
  textFilters: Record<string, string>,
  numericFilters: Record<string, NumericFilterInput>,
): boolean {
  if (column.kind === "text") {
    return matchesTextFilter(column.get(part), textFilters[column.key] ?? "");
  }
  if (column.kind === "options") {
    return matchesOptionFilter(column.get(part), textFilters[column.key] ?? "");
  }
  return matchesNumericFilter(column.get(part), parseRange(numericFilters[column.key]));
}

function isColumnFiltered(
  column: ReportColumn,
  textFilters: Record<string, string>,
  numericFilters: Record<string, NumericFilterInput>,
): boolean {
  if (column.kind === "numeric") {
    const raw = numericFilters[column.key];
    return Boolean(raw && (raw.min.trim() !== "" || raw.max.trim() !== ""));
  }
  return Boolean(textFilters[column.key]?.trim());
}

const inputClass =
  "h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100";
const labelClass = "mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500";

function ColumnFilterControl({
  column,
  textValue,
  onTextChange,
  numericValue,
  onNumericChange,
  options,
}: {
  column: ReportColumn;
  textValue: string;
  onTextChange: (value: string) => void;
  numericValue?: NumericFilterInput;
  onNumericChange?: (value: NumericFilterInput) => void;
  options?: string[];
}) {
  if (column.kind === "numeric") {
    const range = numericValue ?? { min: "", max: "" };
    return (
      <div>
        <span className={labelClass}>{column.label}</span>
        <div className="flex items-center gap-2">
          <input
            type="number"
            placeholder="Min"
            value={range.min}
            onChange={(event) => onNumericChange?.({ ...range, min: event.target.value })}
            className={inputClass}
          />
          <span className="text-stone-400">–</span>
          <input
            type="number"
            placeholder="Max"
            value={range.max}
            onChange={(event) => onNumericChange?.({ ...range, max: event.target.value })}
            className={inputClass}
          />
        </div>
      </div>
    );
  }

  if (column.kind === "options") {
    return (
      <label className="block">
        <span className={labelClass}>{column.label}</span>
        <select value={textValue} onChange={(event) => onTextChange(event.target.value)} className={inputClass}>
          <option value="">Any</option>
          {(options ?? []).map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
      </label>
    );
  }

  return (
    <label className="block">
      <span className={labelClass}>{column.label}</span>
      <input
        type="text"
        value={textValue}
        placeholder={`Search ${column.label.toLowerCase()}…`}
        onChange={(event) => onTextChange(event.target.value)}
        className={inputClass}
      />
    </label>
  );
}

const UPN_HEADER_PATTERN = /^upn\s*#?$/i;

export function ReportsView() {
  const { parts, sourceLabels } = useInventory();
  const [textFilters, setTextFilters] = useState<Record<string, string>>({});
  const [numericFilters, setNumericFilters] = useState<Record<string, NumericFilterInput>>({});

  const standardColumns: ReportColumn[] = useMemo(
    () => [
      { key: "part_number", label: "Part number", kind: "text", alwaysVisible: true, get: (p) => p.part_number },
      { key: "description", label: "Description", kind: "text", alwaysVisible: true, get: (p) => p.description },
      {
        key: "category",
        label: "Category",
        kind: "options",
        alwaysVisible: true,
        get: (p) => (p.category ? [p.category] : []),
      },
      {
        key: "bin_location",
        label: "Bin location",
        kind: "options",
        alwaysVisible: true,
        get: (p) => (p.bin_location ? [p.bin_location] : []),
      },
      {
        key: "quantity_on_hand",
        label: "Quantity on hand",
        kind: "numeric",
        alwaysVisible: true,
        get: (p) => p.quantity_on_hand,
      },
      { key: "status", label: "Status", kind: "options", alwaysVisible: true, get: (p) => [p.status] },
      {
        key: "source",
        label: "Source",
        kind: "options",
        alwaysVisible: true,
        get: (p) => (p.sourceIds ?? []).map((id) => sourceLabels[id] ?? id),
      },
      { key: "tags", label: "Tags", kind: "options", alwaysVisible: false, get: (p) => p.tags ?? [] },
      {
        key: "upn",
        label: "UPN#",
        kind: "text",
        alwaysVisible: false,
        get: (p) => upnFromExtraFields(p.extraFields) ?? "",
      },
    ],
    [sourceLabels],
  );

  // Every other column a source's sheet happens to have (Condition,
  // "Ebay Ready (Yes/No)", "Entry Date", "UPC#", ...) — UPN# is excluded
  // here since it already has its own dedicated column above.
  const dynamicColumns: ReportColumn[] = useMemo(() => {
    return discoverExtraColumns(parts)
      .filter((column) => !UPN_HEADER_PATTERN.test(column.header.trim()))
      .map((column): ReportColumn => {
        const key = `extra:${column.header}`;
        const getRaw = (part: Part): string | undefined => {
          const field = part.extraFields?.[column.header];
          if (!field) return undefined;
          if (field.kind === "boolean") return field.value ? "Yes" : "No";
          return typeof field.value === "string" ? field.value : String(field.value);
        };
        if (column.kind === "boolean" || column.kind === "select") {
          return {
            key,
            label: column.header,
            kind: "options",
            alwaysVisible: false,
            get: (part) => {
              const value = getRaw(part);
              return value ? [value] : [];
            },
          };
        }
        return {
          key,
          label: column.header,
          kind: "text",
          alwaysVisible: false,
          get: (part) => getRaw(part) ?? "",
        };
      });
  }, [parts]);

  const allColumns = useMemo(() => [...standardColumns, ...dynamicColumns], [standardColumns, dynamicColumns]);
  const primaryColumns = useMemo(() => standardColumns.filter((c) => c.alwaysVisible), [standardColumns]);
  const secondaryColumns = useMemo(
    () => [...standardColumns.filter((c) => !c.alwaysVisible), ...dynamicColumns],
    [standardColumns, dynamicColumns],
  );

  const filteredParts = useMemo(
    () => parts.filter((part) => allColumns.every((column) => matchesColumn(column, part, textFilters, numericFilters))),
    [parts, allColumns, textFilters, numericFilters],
  );

  const visibleColumns = useMemo(
    () => allColumns.filter((column) => column.alwaysVisible || isColumnFiltered(column, textFilters, numericFilters)),
    [allColumns, textFilters, numericFilters],
  );

  const hasActiveFilters = allColumns.some((column) => isColumnFiltered(column, textFilters, numericFilters));

  function optionsFor(column: ReportColumn & { kind: "options" }): string[] {
    if (column.key.startsWith("extra:")) {
      return extraColumnValues(parts, column.label);
    }
    return distinctValues(parts.flatMap((part) => column.get(part)));
  }

  function renderControl(column: ReportColumn) {
    return (
      <ColumnFilterControl
        key={column.key}
        column={column}
        textValue={textFilters[column.key] ?? ""}
        onTextChange={(value) => setTextFilters((current) => ({ ...current, [column.key]: value }))}
        numericValue={numericFilters[column.key]}
        onNumericChange={(value) => setNumericFilters((current) => ({ ...current, [column.key]: value }))}
        options={column.kind === "options" ? optionsFor(column) : undefined}
      />
    );
  }

  function handleClearFilters() {
    setTextFilters({});
    setNumericFilters({});
  }

  function handleDownloadCsv() {
    const headers = visibleColumns.map((column) => column.label);
    const rows = filteredParts.map((part) => visibleColumns.map((column) => displayValue(column, part)));
    const csv = toCsv(headers, rows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `parts-report-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="space-y-6">
      <AppHeader
        title="Reports"
        subtitle="Filter the catalog by any column and combine as many filters as you need — every filter narrows the same list."
      />

      <div className="no-print space-y-4">
        <section className="space-y-3 rounded-2xl border border-stone-200 bg-white p-4">
          <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">Standard fields</h2>
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">{primaryColumns.map(renderControl)}</div>
        </section>

        {secondaryColumns.length > 0 ? (
          <details className="rounded-2xl border border-stone-200 bg-white p-4">
            <summary className="cursor-pointer text-sm font-semibold text-stone-700">
              More filters (Tags, UPN#, and sheet-specific columns)
            </summary>
            <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2">{secondaryColumns.map(renderControl)}</div>
          </details>
        ) : null}

        <div className="flex flex-wrap items-center justify-between gap-3">
          <p className="text-sm text-stone-600">
            Showing <span className="font-semibold text-stone-900">{filteredParts.length}</span> of{" "}
            {parts.length} parts
            {hasActiveFilters ? (
              <button
                type="button"
                onClick={handleClearFilters}
                className="ml-2 text-xs font-semibold text-amber-700 underline underline-offset-2"
              >
                Clear filters
              </button>
            ) : null}
          </p>
          <div className="flex gap-2">
            <button
              type="button"
              onClick={handleDownloadCsv}
              disabled={filteredParts.length === 0}
              className="inline-flex items-center gap-1.5 rounded-full bg-stone-900 px-3.5 py-2 text-xs font-semibold text-white disabled:opacity-40"
            >
              <Download className="size-3.5" />
              Download CSV
            </button>
            <button
              type="button"
              onClick={() => window.print()}
              disabled={filteredParts.length === 0}
              className="inline-flex items-center gap-1.5 rounded-full border border-stone-300 bg-white px-3.5 py-2 text-xs font-semibold text-stone-700 disabled:opacity-40"
            >
              <Printer className="size-3.5" />
              Print
            </button>
          </div>
        </div>
      </div>

      <div className="overflow-x-auto rounded-2xl border border-stone-200 bg-white">
        <table className="w-full min-w-max text-left text-sm">
          <thead>
            <tr className="border-b border-stone-200 text-xs font-semibold uppercase tracking-wide text-stone-500">
              {visibleColumns.map((column) => (
                <th key={column.key} className="whitespace-nowrap px-3 py-2">
                  {column.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredParts.length === 0 ? (
              <tr>
                <td colSpan={visibleColumns.length} className="px-3 py-6 text-center text-sm text-stone-500">
                  No parts match these filters.
                </td>
              </tr>
            ) : (
              filteredParts.map((part) => (
                <tr key={part.id} className="border-b border-stone-100 last:border-0">
                  {visibleColumns.map((column) => (
                    <td key={column.key} className="whitespace-nowrap px-3 py-2 text-stone-700">
                      {displayValue(column, part)}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </main>
  );
}
