import { distinctValues } from "./inventory";
import type { ExtraFieldKind, Part } from "./types";

/**
 * Support code for the Reports page (components/ReportsView.tsx) — a
 * generic filter-by-any-column report over the merged parts catalog,
 * including whatever extra columns each source's own sheet happens to
 * have (Condition, "Ebay Ready (Yes/No)", "Entry Date", and so on —
 * see RawPart.extraFields). Kept separate from the component so the
 * trickiest bits (which kind a sheet-specific column should filter as,
 * and CSV escaping) are unit-testable without rendering anything.
 */

export interface DiscoveredExtraColumn {
  header: string;
  kind: ExtraFieldKind;
}

/**
 * Every distinct extra-field header seen anywhere in the catalog, each
 * paired with the kind ("text" | "boolean" | "date" | "select") it's
 * most often classified as. Different sources can disagree on a
 * shared header's kind (the same "Notes" column might be free text in
 * one sheet and happen to look boolean in another, smaller one — see
 * classifyExtraColumn in sharepoint-excel-source.ts, which classifies
 * per source) — this picks whichever kind was seen most often for that
 * header, so one outlier source doesn't flip an otherwise-consistent
 * column's filter type. Sorted alphabetically for a stable column
 * order in the report.
 */
export function discoverExtraColumns(parts: Part[]): DiscoveredExtraColumn[] {
  const countsByHeader = new Map<string, Map<ExtraFieldKind, number>>();
  for (const part of parts) {
    for (const [header, field] of Object.entries(part.extraFields ?? {})) {
      const counts = countsByHeader.get(header) ?? new Map<ExtraFieldKind, number>();
      counts.set(field.kind, (counts.get(field.kind) ?? 0) + 1);
      countsByHeader.set(header, counts);
    }
  }

  return Array.from(countsByHeader.entries())
    .map(([header, counts]) => {
      const [kind] = Array.from(counts.entries()).sort((a, b) => b[1] - a[1])[0];
      return { header, kind };
    })
    .sort((a, b) => a.header.localeCompare(b.header));
}

/**
 * Distinct values one extra-field header actually holds across the
 * catalog, for that column's filter dropdown — boolean cells become
 * "Yes"/"No" so a boolean column's dropdown reads the same way a
 * select/text column's does, rather than needing its own UI. A part
 * with no such column (a different source's sheet) is simply skipped,
 * not counted as a blank value.
 */
export function extraColumnValues(parts: Part[], header: string): string[] {
  const values = parts.map((part) => {
    const field = part.extraFields?.[header];
    if (!field) return undefined;
    if (field.kind === "boolean") return field.value ? "Yes" : "No";
    return typeof field.value === "string" ? field.value : String(field.value);
  });
  return distinctValues(values);
}

/** Whether a text value contains the (trimmed, case-insensitive) query
 *  — an empty query always matches, same as every other filter kind
 *  treating "nothing chosen" as "don't filter on this column". */
export function matchesTextFilter(value: string | undefined, query: string): boolean {
  const trimmed = query.trim();
  if (!trimmed) return true;
  return (value ?? "").toLowerCase().includes(trimmed.toLowerCase());
}

/** Whether a (possibly multi-valued, e.g. Source or Tags) column
 *  includes the exact selected option. An empty selection ("Any")
 *  always matches. */
export function matchesOptionFilter(values: string[], selected: string): boolean {
  if (!selected) return true;
  return values.includes(selected);
}

export interface NumericRange {
  min?: number;
  max?: number;
}

/** Whether a numeric value falls within an (optionally one-sided)
 *  range. Both bounds unset always matches. */
export function matchesNumericFilter(value: number, range: NumericRange): boolean {
  if (range.min !== undefined && value < range.min) return false;
  if (range.max !== undefined && value > range.max) return false;
  return true;
}

/**
 * Renders rows as CSV text (CRLF line endings, per RFC 4180). Any cell
 * containing a comma, double quote, or newline is wrapped in quotes
 * with internal quotes doubled — the two things that silently corrupt
 * a naive `values.join(",")` export (a bin location or description
 * that happens to contain a comma would otherwise shift every later
 * column over by one).
 */
export function toCsv(headers: string[], rows: string[][]): string {
  const escapeCell = (cell: string): string => {
    if (/[",\r\n]/.test(cell)) {
      return `"${cell.replace(/"/g, '""')}"`;
    }
    return cell;
  };
  return [headers, ...rows]
    .map((row) => row.map(escapeCell).join(","))
    .join("\r\n");
}
