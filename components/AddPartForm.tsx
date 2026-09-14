"use client";

import { AlertTriangle, ArrowLeft, Hash, PackagePlus } from "lucide-react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useEffect, useMemo, useRef, useState } from "react";
import { FormField } from "@/components/FormField";
import { useInventory } from "@/components/InventoryProvider";
import { distinctValues } from "@/lib/inventory";
import { CONDITION_OPTIONS, type ExtraFields } from "@/lib/types";

/** The sheet header this app's Condition dropdown writes to — matches
 *  the "select" field name sharepoint-excel-source.ts classifies (see
 *  SELECT_FIELD_OPTIONS there). Kept as one literal here since, unlike
 *  UPN#, this app doesn't know a source's UPN header ahead of time —
 *  see the "Next UPN#" preview below, which learns it from
 *  /api/parts/next-upn instead. */
const CONDITION_HEADER = "Condition";

/** The sheet header this app's Ebay-ready dropdown writes to — matches
 *  the exact "(Yes/No)" spelling classifyExtraColumn's headerNamesBoolean
 *  looks for in sharepoint-excel-source.ts, so a source with this
 *  column reads the value back the same way it was written. A source
 *  without this column simply ignores the field (see fieldAssignments),
 *  same as Condition. */
const EBAY_READY_HEADER = "Ebay Ready (Yes/No)";

/** Sentinel <option> value for "none of the below — let me type one",
 *  distinct from "" (which means "no location picked at all") so the
 *  bin location <select> can tell the two apart. */
const ADD_NEW_LOCATION = "__add-new-location__";

/** Which <option> the bin location <select> should show as selected
 *  for the current form value: the value itself when it's one of this
 *  source's known locations, "" when the field is simply empty, or the
 *  "Add a new location…" sentinel for anything else (a location typed
 *  in before switching to a source that doesn't have it yet, say). */
function binLocationSelectValue(value: string, options: string[]): string {
  const trimmed = value.trim();
  if (trimmed === "") return "";
  return options.includes(trimmed) ? trimmed : ADD_NEW_LOCATION;
}

export function AddPartForm() {
  const router = useRouter();
  const { parts, writableSources, applyNewPart } = useInventory();
  const [form, setForm] = useState({
    sourceId: writableSources[0]?.id ?? "",
    part_number: "",
    description: "",
    bin_location: "",
    quantity_on_hand: "0",
    condition: "",
    ebayReady: "",
    upn: "",
  });
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Every bin location already in use by a part that touches the
  // currently selected source — a fresh part almost always belongs
  // wherever other parts from the same source already sit, so offering
  // these as a dropdown beats retyping a bin string from memory (and
  // the inevitable typo variants that come with it: "A-12-04" vs
  // "A1204" vs "a-12-4"). Scoped to the selected source, not the whole
  // catalog, since a bin string valid in one source's sheet isn't
  // necessarily meaningful in another's. A part merged from more than
  // one source (see mergeParts) only has the one bin_location that won
  // the merge, so this can occasionally offer a location that actually
  // came from a different source that also touches this part — a minor
  // approximation, not a source of bad data (it's just a suggestion the
  // technician picks from, never written back unreviewed).
  const binLocationOptions = useMemo(
    () =>
      distinctValues(
        parts
          .filter((part) => (part.sourceIds ?? []).includes(form.sourceId))
          .map((part) => part.bin_location),
      ),
    [parts, form.sourceId],
  );

  // The exact UPN-shaped header this source's sheet actually uses (e.g.
  // "UPN#" vs "UPN #"), learned from the preview fetch below — needed
  // so submit sends the value under the header the sheet expects rather
  // than a guess. Null means "this source has no UPN-shaped column" (or
  // the preview hasn't resolved yet), in which case no UPN field shows
  // at all and the source's own auto-numbering (or lack of it) at save
  // time is unaffected.
  const [upnHeader, setUpnHeader] = useState<string | null>(null);
  const [upnLoading, setUpnLoading] = useState(false);
  // Whether the technician has hand-edited the picked number — a ref
  // (not state) so the effect below can check it without needing to
  // re-run every time it changes, which would trigger a pointless
  // refetch on every keystroke.
  const upnTouchedRef = useRef(false);

  // Picks the next free UPN# as soon as a source is selected — right
  // when the form opens, using the default source, and again whenever
  // the technician switches sources — rather than only learning the
  // number after the part is already saved (see peekNextUpn).
  useEffect(() => {
    const sourceId = form.sourceId;
    // A source is only ever missing before any writable source loads
    // in (the form itself doesn't render at all in that case — see the
    // "No writable source yet" branch below) — nothing to preview, and
    // nothing was showing a preview yet either, so there's no state to
    // reset here.
    if (!sourceId) {
      return;
    }
    let cancelled = false;
    upnTouchedRef.current = false;

    // Wrapped in its own async function (rather than setState calls
    // sitting directly in the effect body) purely to satisfy the
    // set-state-in-effect lint rule — the actual behavior is the usual
    // "fetch in an effect" shape: flip on a loading flag, fetch, then
    // apply whatever came back unless this run was superseded.
    async function loadUpnPreview() {
      setUpnLoading(true);
      try {
        const response = await fetch(`/api/parts/next-upn?sourceId=${encodeURIComponent(sourceId)}`);
        const payload: { header?: string | null; next?: string } = response.ok
          ? await response.json()
          : { header: null };
        if (cancelled) return;
        setUpnHeader(payload.header ?? null);
        if (payload.header && payload.next && !upnTouchedRef.current) {
          setForm((f) => (f.sourceId === sourceId ? { ...f, upn: payload.next as string } : f));
        }
      } catch {
        if (!cancelled) setUpnHeader(null);
      } finally {
        if (!cancelled) setUpnLoading(false);
      }
    }

    loadUpnPreview();
    return () => {
      cancelled = true;
    };
  }, [form.sourceId]);

  if (writableSources.length === 0) {
    return (
      <main className="space-y-4">
        <Link
          href="/"
          className="inline-flex items-center gap-1.5 text-sm font-semibold text-stone-600"
        >
          <ArrowLeft className="size-4" />
          All parts
        </Link>
        <section className="rounded-2xl border border-dashed border-stone-300 bg-white p-4">
          <h2 className="text-sm font-semibold text-stone-900">No writable source yet</h2>
          <p className="mt-1.5 text-sm leading-6 text-stone-600">
            Adding a part needs a SharePoint workbook to add it to. Add one from{" "}
            <Link href="/sources" className="font-semibold text-amber-700 underline">
              Data sources
            </Link>{" "}
            first.
          </p>
        </section>
      </main>
    );
  }

  async function handleSubmit(event: React.FormEvent) {
    event.preventDefault();
    setSubmitting(true);
    setError(null);
    try {
      const condition = form.condition.trim();
      const upnValue = form.upn.trim();
      const extraFields: ExtraFields = {};
      if (condition) {
        extraFields[CONDITION_HEADER] = {
          kind: "select",
          value: condition,
          options: [...CONDITION_OPTIONS],
        };
      }
      // Sends the picked-at-form-open number under this source's actual
      // UPN header (learned from /api/parts/next-upn) rather than
      // guessing a column name — withAutoNextUpn on the server only
      // fills a blank UPN#, so a value supplied here always wins,
      // matching what the technician saw on screen instead of
      // whatever's freshest at save time.
      if (upnHeader && upnValue) {
        extraFields[upnHeader] = { kind: "text", value: upnValue };
      }
      if (form.ebayReady) {
        extraFields[EBAY_READY_HEADER] = { kind: "boolean", value: form.ebayReady === "yes" };
      }
      const response = await fetch("/api/parts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sourceId: form.sourceId,
          part_number: form.part_number,
          description: form.description,
          bin_location: form.bin_location,
          quantity_on_hand: Number(form.quantity_on_hand),
          extraFields,
        }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to add part.");
      }
      const newId = applyNewPart(
        {
          part_number: form.part_number.trim(),
          description: form.description.trim(),
          bin_location: form.bin_location.trim(),
          quantity_on_hand: Math.max(0, Number(form.quantity_on_hand)),
          extraFields,
        },
        form.sourceId,
      );
      router.push(`/parts/${newId}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to add part.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <main className="space-y-6">
      <Link
        href="/"
        className="inline-flex items-center gap-1.5 text-sm font-semibold text-stone-600"
      >
        <ArrowLeft className="size-4" />
        All parts
      </Link>

      <div className="flex items-center gap-3">
        <div className="flex size-11 shrink-0 items-center justify-center rounded-xl bg-amber-50 text-amber-700">
          <PackagePlus className="size-5" />
        </div>
        <h1 className="text-lg font-semibold text-stone-900">Add a part</h1>
      </div>

      {error ? (
        <p className="inline-flex items-start gap-2 rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-3 text-sm leading-5 text-rose-900">
          <AlertTriangle className="mt-0.5 size-4 shrink-0" />
          {error}
        </p>
      ) : null}

      <form onSubmit={handleSubmit} className="space-y-3 rounded-2xl border border-stone-200 bg-white p-4">
        {upnHeader ? (
          <div className="rounded-xl border border-amber-200 bg-amber-50 p-3">
            <span className="mb-1 block text-[11px] font-semibold uppercase tracking-wide text-amber-700">
              Next {upnHeader}
            </span>
            <div className="flex items-center gap-2">
              <Hash className="size-4 shrink-0 text-amber-700" />
              <input
                value={form.upn}
                onChange={(event) => {
                  upnTouchedRef.current = true;
                  setForm((f) => ({ ...f, upn: event.target.value }));
                }}
                className="h-9 w-28 rounded-lg border border-amber-300 bg-white px-2 font-mono text-sm outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-100"
              />
              <span className="text-xs leading-4 text-amber-700">
                Picked automatically — change it if this one shouldn&apos;t be used.
              </span>
            </div>
          </div>
        ) : upnLoading ? (
          <p className="text-xs text-stone-400">Checking for a UPN# column…</p>
        ) : null}

        <label className="block">
          <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
            Add to source
          </span>
          <select
            required
            value={form.sourceId}
            onChange={(event) => setForm((f) => ({ ...f, sourceId: event.target.value }))}
            className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
          >
            {writableSources.map((source) => (
              <option key={source.id} value={source.id}>
                {source.label}
              </option>
            ))}
          </select>
        </label>

        <FormField
          label="Part number"
          placeholder="WR17X11705"
          value={form.part_number}
          onChange={(value) => setForm((f) => ({ ...f, part_number: value }))}
        />
        <FormField
          label="Description"
          placeholder="Water filter cartridge"
          value={form.description}
          onChange={(value) => setForm((f) => ({ ...f, description: value }))}
          required={false}
        />
        {binLocationOptions.length > 0 ? (
          <div>
            <label className="block">
              <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
                Bin location (optional)
              </span>
              <select
                value={binLocationSelectValue(form.bin_location, binLocationOptions)}
                onChange={(event) => {
                  const next = event.target.value;
                  setForm((f) => ({
                    ...f,
                    bin_location: next === ADD_NEW_LOCATION ? "" : next,
                  }));
                }}
                className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
              >
                <option value="">— None —</option>
                {binLocationOptions.map((location) => (
                  <option key={location} value={location}>
                    {location}
                  </option>
                ))}
                <option value={ADD_NEW_LOCATION}>Add a new location…</option>
              </select>
            </label>
            {binLocationSelectValue(form.bin_location, binLocationOptions) === ADD_NEW_LOCATION ? (
              <input
                autoFocus
                placeholder="A-12-04"
                value={form.bin_location}
                onChange={(event) =>
                  setForm((f) => ({ ...f, bin_location: event.target.value }))
                }
                className="mt-2 h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
              />
            ) : null}
          </div>
        ) : (
          <FormField
            label="Bin location"
            placeholder="A-12-04"
            value={form.bin_location}
            onChange={(value) => setForm((f) => ({ ...f, bin_location: value }))}
            required={false}
          />
        )}
        <FormField
          label="Quantity on hand"
          type="number"
          value={form.quantity_on_hand}
          onChange={(value) => setForm((f) => ({ ...f, quantity_on_hand: value }))}
        />

        <label className="block">
          <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
            Condition
          </span>
          <select
            value={form.condition}
            onChange={(event) => setForm((f) => ({ ...f, condition: event.target.value }))}
            className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
          >
            <option value="">— Not set —</option>
            {CONDITION_OPTIONS.filter((option) => option !== "Other").map((option) => (
              <option key={option} value={option}>
                {option}
              </option>
            ))}
          </select>
        </label>

        <label className="block">
          <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
            Ebay ready
          </span>
          <select
            value={form.ebayReady}
            onChange={(event) => setForm((f) => ({ ...f, ebayReady: event.target.value }))}
            className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
          >
            <option value="">— Not set —</option>
            <option value="yes">Yes</option>
            <option value="no">No</option>
          </select>
        </label>

        <button
          type="submit"
          disabled={submitting}
          className="w-full rounded-xl bg-stone-900 px-3 py-2.5 text-sm font-semibold text-white disabled:opacity-60"
        >
          {submitting ? "Adding…" : "Add part to SharePoint"}
        </button>
      </form>
    </main>
  );
}
