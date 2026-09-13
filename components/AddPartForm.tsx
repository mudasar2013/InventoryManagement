"use client";

import { AlertTriangle, ArrowLeft, PackagePlus } from "lucide-react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { FormField } from "@/components/FormField";
import { useInventory } from "@/components/InventoryProvider";
import { CONDITION_OPTIONS } from "@/lib/types";

/** The sheet header this app's Condition dropdown writes to — matches
 *  the "select" field name sharepoint-excel-source.ts classifies (see
 *  SELECT_FIELD_OPTIONS there). Kept as one literal here since, unlike
 *  UPN#, this is the only extra field the Add form sets by name. */
const CONDITION_HEADER = "Condition";

export function AddPartForm() {
  const router = useRouter();
  const { writableSources, applyNewPart } = useInventory();
  const [form, setForm] = useState({
    sourceId: writableSources[0]?.id ?? "",
    part_number: "",
    description: "",
    bin_location: "",
    quantity_on_hand: "0",
    condition: "",
  });
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

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
      const extraFields = condition
        ? { [CONDITION_HEADER]: { kind: "select" as const, value: condition, options: [...CONDITION_OPTIONS] } }
        : undefined;
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
        <FormField
          label="Bin location"
          placeholder="A-12-04"
          value={form.bin_location}
          onChange={(value) => setForm((f) => ({ ...f, bin_location: value }))}
          required={false}
        />
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
