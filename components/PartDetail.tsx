"use client";

import {
  AlertTriangle,
  ArrowLeft,
  Check,
  ClipboardList,
  MapPin,
  Minus,
  Pencil,
  Tag,
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
import type { ExtraFields } from "@/lib/types";

export function PartDetail({ partId }: { partId: string }) {
  const router = useRouter();
  const { parts, jobsForPart, writableSources, applyPartUpdate, applyBulkUpdate } =
    useInventory();
  const part = parts.find((item) => item.id === partId);
  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState({
    part_number: "",
    description: "",
    bin_location: "",
    quantity_on_hand: "0",
    category: "",
    sourceId: "",
    extraFields: {} as ExtraFields,
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
  const extraFieldEntries = Object.entries(part.extraFields ?? {});

  function startEditing() {
    setError(null);
    if (!part) return;
    setForm({
      part_number: part.part_number,
      description: part.description,
      bin_location: part.bin_location,
      quantity_on_hand: String(part.quantity_on_hand),
      category: part.category ?? "",
      sourceId: editableSources[0]?.id ?? "",
      extraFields: part.extraFields ? { ...part.extraFields } : {},
    });
    setEditing(true);
  }

  function setExtraFieldValue(header: string, value: string | boolean) {
    setForm((f) => ({
      ...f,
      extraFields: {
        ...f.extraFields,
        [header]: { ...f.extraFields[header], value },
      },
    }));
  }

  /** Formats a "date" extra field's normalized "YYYY-MM-DD" value for
   *  the read-only display — same value the edit form's
   *  <input type="date"> uses, just shown the way a person reads a
   *  date rather than the ISO form a date input requires. */
  function formatExtraDate(value: string | boolean): string {
    const iso = typeof value === "string" ? value.trim() : "";
    if (!iso) return "—";
    const parsed = new Date(`${iso}T00:00:00`);
    return Number.isNaN(parsed.getTime()) ? iso : parsed.toLocaleDateString();
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
          category: form.category,
          extraFields: form.extraFields,
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
        category: form.category.trim(),
        extraFields: form.extraFields,
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
            <Tag className="size-6" />
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
        {part.category ? (
          <p className="mt-2 inline-flex items-center rounded-full bg-stone-100 px-2.5 py-1 text-xs font-semibold text-stone-600">
            {part.category}
          </p>
        ) : null}
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
            <FormField
              label="Category"
              value={form.category}
              onChange={(value) => setForm((f) => ({ ...f, category: value }))}
              required={false}
            />

            {Object.keys(form.extraFields).length > 0 ? (
              <div className="space-y-3 border-t border-stone-100 pt-3">
                <p className="text-xs font-semibold uppercase tracking-wide text-stone-500">
                  Every other column this source has
                </p>
                {Object.entries(form.extraFields).map(([header, field]) =>
                  field.kind === "boolean" ? (
                    <label key={header} className="flex items-center justify-between gap-2">
                      <span className="text-sm text-stone-700">{header}</span>
                      <select
                        value={field.value ? "yes" : "no"}
                        onChange={(event) => setExtraFieldValue(header, event.target.value === "yes")}
                        className="h-9 rounded-lg border border-stone-200 bg-white px-2 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
                      >
                        <option value="no">No</option>
                        <option value="yes">Yes</option>
                      </select>
                    </label>
                  ) : field.kind === "date" ? (
                    <label key={header} className="block">
                      <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
                        {header}
                      </span>
                      <input
                        type="date"
                        value={typeof field.value === "string" ? field.value : ""}
                        onChange={(event) => setExtraFieldValue(header, event.target.value)}
                        className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
                      />
                    </label>
                  ) : field.kind === "select" ? (
                    <SelectOrOtherField
                      key={header}
                      label={header}
                      value={typeof field.value === "string" ? field.value : ""}
                      options={field.options ?? []}
                      onChange={(value) => setExtraFieldValue(header, value)}
                    />
                  ) : (
                    <FormField
                      key={header}
                      label={header}
                      value={String(field.value)}
                      onChange={(value) => setExtraFieldValue(header, value)}
                      required={false}
                    />
                  ),
                )}
              </div>
            ) : null}

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

      <TagsEditor part={part} applyBulkUpdate={applyBulkUpdate} />

      {extraFieldEntries.length > 0 ? (
        <section className="space-y-2">
          <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">
            More from the source sheet
          </h2>
          <div className="rounded-2xl border border-stone-200 bg-white p-4">
            <dl className="grid grid-cols-2 gap-3 text-xs">
              {extraFieldEntries.map(([header, field]) => (
                <div key={header}>
                  <dt className="font-semibold uppercase tracking-wide text-stone-500">
                    {header}
                  </dt>
                  <dd className="mt-0.5 font-medium text-stone-800">
                    {field.kind === "boolean" ? (
                      field.value ? (
                        <span className="inline-flex items-center gap-1 text-emerald-700">
                          <Check className="size-3.5" /> Yes
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 text-stone-400">
                          <Minus className="size-3.5" /> No
                        </span>
                      )
                    ) : field.kind === "date" ? (
                      formatExtraDate(field.value)
                    ) : (
                      String(field.value) || "—"
                    )}
                  </dd>
                </div>
              ))}
            </dl>
          </div>
        </section>
      ) : null}

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

/** A dropdown for a "select"-kind extra field (currently just
 *  "Condition" — see SELECT_FIELD_OPTIONS in sharepoint-excel-source.ts)
 *  with a graceful escape hatch for a value the sheet already has that
 *  isn't in the fixed list: rather than silently discarding or
 *  overwriting it, the dropdown shows "Other" selected and a free-text
 *  input pre-filled with the actual value, so it stays visible and
 *  editable instead of forcing it into one of the fixed choices. */
function SelectOrOtherField({
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
  const trimmed = value.trim();
  const knownOptions = options.filter((option) => option.toLowerCase() !== "other");
  const matched = knownOptions.find((option) => option.toLowerCase() === trimmed.toLowerCase());
  const selectValue = trimmed === "" ? "" : (matched ?? "Other");

  return (
    <div>
      <label className="block">
        <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
          {label}
        </span>
        <select
          value={selectValue}
          onChange={(event) => {
            const next = event.target.value;
            onChange(next === "Other" ? "" : next);
          }}
          className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
        >
          <option value="">—</option>
          {options.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
      </label>
      {selectValue === "Other" ? (
        <input
          value={value}
          onChange={(event) => onChange(event.target.value)}
          placeholder={`Specify ${label.toLowerCase()}`}
          className="mt-2 h-10 w-full rounded-lg border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
        />
      ) : null}
    </div>
  );
}

/** The tags card + its own small edit affordance — kept independent of
 *  the main "Edit this part" form above because tags are app-managed
 *  (see lib/tags-store.ts) and apply to a part regardless of whether
 *  it has any writable SharePoint source at all. A part from the local
 *  demo catalog can be tagged even though it can never be "edited" in
 *  the SharePoint sense. */
function TagsEditor({
  part,
  applyBulkUpdate,
}: {
  part: { id: string; part_number: string; tags?: string[] };
  applyBulkUpdate: (updates: { part_number: string; fields: { tags?: string[] } }[]) => void;
}) {
  const { tags: vocabulary } = useInventory();
  const [editing, setEditing] = useState(false);
  const [selected, setSelected] = useState<string[]>([]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  function startEditing() {
    setError(null);
    setSelected(part.tags ?? []);
    setEditing(true);
  }

  function toggle(tag: string) {
    setSelected((current) =>
      current.includes(tag) ? current.filter((t) => t !== tag) : [...current, tag],
    );
  }

  async function handleSave() {
    setSaving(true);
    setError(null);
    try {
      const response = await fetch("/api/parts/tags", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ part_number: part.part_number, tags: selected }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to update tags.");
      }
      applyBulkUpdate([{ part_number: part.part_number, fields: { tags: selected } }]);
      setEditing(false);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to update tags.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <section className="space-y-2">
      <div className="flex items-center justify-between">
        <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">Tags</h2>
        {!editing ? (
          <button
            type="button"
            onClick={startEditing}
            aria-label="Edit tags"
            className="rounded-full p-1.5 text-stone-400 hover:bg-amber-50 hover:text-amber-700"
          >
            <Pencil className="size-3.5" />
          </button>
        ) : null}
      </div>

      {error ? (
        <p className="inline-flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2.5 text-sm leading-5 text-rose-900">
          <AlertTriangle className="mt-0.5 size-4 shrink-0" />
          {error}
        </p>
      ) : null}

      {editing ? (
        vocabulary.length === 0 ? (
          <p className="rounded-2xl border border-dashed border-stone-300 bg-white px-4 py-4 text-sm text-stone-500">
            No tags exist yet — add some from Settings first.
          </p>
        ) : (
          <div className="space-y-3 rounded-2xl border border-stone-200 bg-white p-4">
            <div className="flex flex-wrap gap-2">
              {vocabulary.map((tag) => {
                const active = selected.includes(tag);
                return (
                  <button
                    key={tag}
                    type="button"
                    onClick={() => toggle(tag)}
                    className={`rounded-full px-3 py-1.5 text-xs font-semibold ring-1 ${
                      active
                        ? "bg-stone-900 text-white ring-stone-900"
                        : "bg-white text-stone-600 ring-stone-200"
                    }`}
                  >
                    {tag}
                  </button>
                );
              })}
            </div>
            <div className="flex gap-2">
              <button
                type="button"
                onClick={() => setEditing(false)}
                className="flex-1 rounded-xl border border-stone-200 px-3 py-2 text-sm font-semibold text-stone-700"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSave}
                disabled={saving}
                className="flex-1 rounded-xl bg-stone-900 px-3 py-2 text-sm font-semibold text-white disabled:opacity-60"
              >
                {saving ? "Saving…" : "Save tags"}
              </button>
            </div>
          </div>
        )
      ) : (part.tags ?? []).length > 0 ? (
        <div className="flex flex-wrap gap-2">
          {(part.tags ?? []).map((tag) => (
            <span
              key={tag}
              className="inline-flex items-center gap-1 rounded-full bg-stone-100 px-2.5 py-1 text-xs font-semibold text-stone-600"
            >
              <Tag className="size-3" />
              {tag}
            </span>
          ))}
        </div>
      ) : (
        <p className="text-sm text-stone-500">No tags yet.</p>
      )}
    </section>
  );
}
