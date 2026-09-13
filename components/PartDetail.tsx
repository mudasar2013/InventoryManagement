"use client";

import {
  AlertTriangle,
  ArrowLeft,
  ClipboardList,
  MapPin,
  Package,
  Pencil,
  X,
} from "lucide-react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { AttachPartToJob } from "@/components/AttachPartToJob";
import { FormField } from "@/components/FormField";
import { useInventory } from "@/components/InventoryProvider";
import { LinkBanner } from "@/components/LinkBanner";
import { JobStatusBadge, PartStatusBadge } from "@/components/StatusBadge";

export function PartDetail({ partId }: { partId: string }) {
  const router = useRouter();
  const { parts, jobsForPart, writableSources, applyPartUpdate } = useInventory();
  const part = parts.find((item) => item.id === partId);
  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState({
    part_number: "",
    description: "",
    bin_location: "",
    quantity_on_hand: "0",
    sourceId: "",
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  if (!part) {
    return (
      <main className="space-y-4">
        <Link
          href="/"
          className="inline-flex items-center gap-1.5 text-sm font-semibold text-stone-600"
        >
          <ArrowLeft className="size-4" />
          All parts
        </Link>
        <p className="text-sm text-stone-600">That part is not in this catalog.</p>
      </main>
    );
  }

  const relatedJobs = jobsForPart(part.id);

  // Only SharePoint sources this part is actually reported by can be
  // written to — "local" (the bundled demo catalog) never appears in
  // writableSources, and a source that used to report this part but
  // was since removed won't either.
  const editableSourceIds = (part.sourceIds ?? []).filter((id) =>
    writableSources.some((source) => source.id === id),
  );
  const editableSources = writableSources.filter((source) =>
    editableSourceIds.includes(source.id),
  );

  function startEditing() {
    setError(null);
    if (!part) return;
    setForm({
      part_number: part.part_number,
      description: part.description,
      bin_location: part.bin_location,
      quantity_on_hand: String(part.quantity_on_hand),
      sourceId: editableSources[0]?.id ?? "",
    });
    setEditing(true);
  }

  async function handleSave(event: React.FormEvent) {
    event.preventDefault();
    if (!part) return;
    setSaving(true);
    setError(null);
    try {
      const response = await fetch("/api/parts", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sourceId: form.sourceId,
          existingPartNumber: part.part_number,
          part_number: form.part_number,
          description: form.description,
          bin_location: form.bin_location,
          quantity_on_hand: Number(form.quantity_on_hand),
        }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to update part.");
      }
      const fields = {
        part_number: form.part_number.trim(),
        description: form.description.trim(),
        bin_location: form.bin_location.trim(),
        quantity_on_hand: Math.max(0, Number(form.quantity_on_hand)),
      };
      const newId = applyPartUpdate(part.id, fields);
      setEditing(false);
      if (newId !== part.id) {
        router.replace(`/parts/${newId}`);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to update part.");
    } finally {
      setSaving(false);
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

      <LinkBanner />

      <section className="rounded-2xl border border-stone-200 bg-white p-5 shadow-sm">
        <div className="flex items-start justify-between gap-3">
          <div className="flex size-12 items-center justify-center rounded-xl bg-amber-50 text-amber-700">
            <Package className="size-6" />
          </div>
          <div className="flex items-center gap-2">
            <PartStatusBadge status={part.status} />
            {editableSources.length > 0 && !editing ? (
              <button
                type="button"
                onClick={startEditing}
                aria-label="Edit part"
                className="rounded-full p-1.5 text-stone-400 hover:bg-amber-50 hover:text-amber-700"
              >
                <Pencil className="size-4" />
              </button>
            ) : null}
          </div>
        </div>
        <p className="mt-3 text-[11px] font-semibold uppercase tracking-wide text-stone-500">
          Part number
        </p>
        <p className="font-mono text-xl font-semibold tracking-wide text-stone-900">
          {part.part_number}
        </p>
        <p className="mt-2 text-base leading-6 text-stone-600">{part.description}</p>
      </section>

      {editing ? (
        <section className="space-y-3 rounded-2xl border border-stone-200 bg-white p-4">
          <div className="flex items-center justify-between gap-2">
            <p className="text-sm font-semibold text-stone-900">Edit this part</p>
            <button
              type="button"
              onClick={() => setEditing(false)}
              className="rounded-full p-1 text-stone-400"
              aria-label="Cancel edit"
            >
              <X className="size-4" />
            </button>
          </div>

          {error ? (
            <p className="inline-flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2.5 text-sm leading-5 text-rose-900">
              <AlertTriangle className="mt-0.5 size-4 shrink-0" />
              {error}
            </p>
          ) : null}

          <form onSubmit={handleSave} className="space-y-3">
            {editableSources.length > 1 ? (
              <label className="block">
                <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
                  This part is reported by more than one source — save to
                </span>
                <select
                  required
                  value={form.sourceId}
                  onChange={(event) => setForm((f) => ({ ...f, sourceId: event.target.value }))}
                  className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
                >
                  {editableSources.map((source) => (
                    <option key={source.id} value={source.id}>
                      {source.label}
                    </option>
                  ))}
                </select>
              </label>
            ) : null}

            <FormField
              label="Part number"
              value={form.part_number}
              onChange={(value) => setForm((f) => ({ ...f, part_number: value }))}
            />
            <FormField
              label="Description"
              value={form.description}
              onChange={(value) => setForm((f) => ({ ...f, description: value }))}
              required={false}
            />
            <FormField
              label="Bin location"
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

            <button
              type="submit"
              disabled={saving}
              className="w-full rounded-xl bg-stone-900 px-3 py-2.5 text-sm font-semibold text-white disabled:opacity-60"
            >
              {saving ? "Saving…" : "Save to SharePoint"}
            </button>
          </form>
        </section>
      ) : null}

      <section className="grid grid-cols-2 gap-3">
        <div className="rounded-2xl border border-stone-200 bg-white p-4">
          <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
            Bin location
          </p>
          <p className="mt-2 inline-flex items-center gap-1.5 text-lg font-semibold text-stone-900">
            <MapPin className="size-4 text-amber-700" />
            {part.bin_location}
          </p>
        </div>
        <div className="rounded-2xl border border-stone-200 bg-white p-4">
          <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
            Quantity on hand
          </p>
          <p className="mt-2 text-lg font-semibold text-stone-900">
            {part.quantity_on_hand}
          </p>
        </div>
      </section>

      <section className="space-y-3">
        <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">
          Jobs using this part
        </h2>
        {relatedJobs.length === 0 ? (
          <div className="rounded-2xl border border-dashed border-stone-300 bg-white px-4 py-8 text-center">
            <ClipboardList className="mx-auto size-7 text-stone-400" />
            <p className="mt-2 text-sm font-medium text-stone-700">
              Not assigned to an open job
            </p>
            <p className="mt-1 text-sm text-stone-500">
              Attach it below. Stock quantity will not change.
            </p>
          </div>
        ) : (
          <ul className="space-y-2">
            {relatedJobs.map((job) => (
              <li key={job.id}>
                <Link
                  href="/jobs"
                  className="block rounded-2xl border border-stone-200 bg-white p-4"
                >
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <p className="font-mono text-xs font-semibold text-stone-500">
                        Job {job.job_number}
                      </p>
                      <p className="mt-1 font-semibold text-stone-900">
                        {job.customer_name}
                      </p>
                      <p className="mt-1 text-sm text-stone-600">{job.appliance}</p>
                    </div>
                    <JobStatusBadge status={job.status} />
                  </div>
                </Link>
              </li>
            ))}
          </ul>
        )}
      </section>

      <AttachPartToJob part={part} />
    </main>
  );
}
