"use client";

import { AlertTriangle, X } from "lucide-react";
import { useState } from "react";
import type { EditablePartFields } from "@/components/InventoryProvider";
import { useInventory } from "@/components/InventoryProvider";
import type { Part } from "@/lib/types";

const NONE = "";

/**
 * Bulk-edits every part passed in `parts` (the Parts page's current
 * selection — see PartsExplorer.tsx). Each field has its own checkbox
 * ("Set X to...") so leaving a field unchecked means exactly that —
 * don't touch it — rather than trying to infer intent from an empty
 * input. Quantity/bin location/category go to each part's own
 * SharePoint source (skipped, per-part, for a part with no writable
 * source — tags still apply to those); tags are app-managed and apply
 * to every selected part regardless of source.
 */
export function BulkEditPanel({
  parts,
  onClose,
  onDone,
}: {
  parts: Part[];
  onClose: () => void;
  onDone: () => void;
}) {
  const { writableSources, tags, applyBulkUpdate } = useInventory();

  const [setQuantity, setSetQuantity] = useState(false);
  const [quantityValue, setQuantityValue] = useState("0");
  const [setBinLocation, setSetBinLocation] = useState(false);
  const [binLocationValue, setBinLocationValue] = useState("");
  const [setCategory, setSetCategory] = useState(false);
  const [categoryValue, setCategoryValue] = useState("");
  const [addTag, setAddTag] = useState(NONE);
  const [removeTag, setRemoveTag] = useState(NONE);

  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const writableSourceIds = new Set(writableSources.map((source) => source.id));
  const editablePartCount = parts.filter((part) =>
    (part.sourceIds ?? []).some((id) => writableSourceIds.has(id)),
  ).length;
  const fieldsSelected = setQuantity || setBinLocation || setCategory;
  const tagsSelected = Boolean(addTag || removeTag);

  async function handleSubmit(event: React.FormEvent) {
    event.preventDefault();
    if (!fieldsSelected && !tagsSelected) return;
    setSubmitting(true);
    setError(null);

    const updates = parts.flatMap((part) => {
      if (!fieldsSelected) return [];
      const sourceId = (part.sourceIds ?? []).find((id) => writableSourceIds.has(id));
      if (!sourceId) return [];
      return [
        {
          sourceId,
          existingPartNumber: part.part_number,
          fields: {
            part_number: part.part_number,
            description: part.description,
            bin_location: setBinLocation ? binLocationValue.trim() : part.bin_location,
            quantity_on_hand: setQuantity
              ? Math.max(0, Number(quantityValue))
              : part.quantity_on_hand,
            category: setCategory ? categoryValue.trim() : part.category,
            extraFields: part.extraFields,
          },
        },
      ];
    });

    try {
      const response = await fetch("/api/parts/bulk", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          updates,
          partNumbers: parts.map((part) => part.part_number),
          addTag: addTag || undefined,
          removeTag: removeTag || undefined,
        }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Bulk edit failed.");
      }

      const results = (payload.results ?? []) as { part_number: string; ok: boolean; error?: string }[];
      const failed = results.filter((result) => !result.ok);

      const localUpdates = parts.map((part) => {
        const fields: Partial<EditablePartFields> = {};
        const succeeded = !fieldsSelected || results.find((r) => r.part_number === part.part_number)?.ok;
        if (succeeded) {
          if (setQuantity) fields.quantity_on_hand = Math.max(0, Number(quantityValue));
          if (setBinLocation) fields.bin_location = binLocationValue.trim();
          if (setCategory) fields.category = categoryValue.trim();
        }
        if (tagsSelected) {
          let nextTags = part.tags ?? [];
          if (addTag && !nextTags.includes(addTag)) nextTags = [...nextTags, addTag];
          if (removeTag) nextTags = nextTags.filter((tag) => tag !== removeTag);
          fields.tags = nextTags;
        }
        return { part_number: part.part_number, fields };
      });
      applyBulkUpdate(localUpdates);

      if (failed.length > 0) {
        setError(
          `${failed.length} of ${updates.length} part${updates.length === 1 ? "" : "s"} failed to save: ` +
            failed.map((f) => `${f.part_number} (${f.error})`).join("; "),
        );
        setSubmitting(false);
        return;
      }

      onDone();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Bulk edit failed.");
      setSubmitting(false);
    }
  }

  return (
    <div className="fixed inset-0 z-30 flex items-end justify-center bg-stone-900/40 sm:items-center">
      <div className="max-h-[85vh] w-full max-w-lg overflow-y-auto rounded-t-2xl bg-white p-5 shadow-xl sm:rounded-2xl">
        <div className="flex items-center justify-between gap-2">
          <div>
            <h2 className="text-base font-semibold text-stone-900">
              Bulk edit {parts.length} part{parts.length === 1 ? "" : "s"}
            </h2>
            {editablePartCount < parts.length ? (
              <p className="mt-1 text-xs text-stone-500">
                {parts.length - editablePartCount} of these have no writable source — quantity,
                location, and category changes will only apply to the other {editablePartCount}.
                Tags still apply to all.
              </p>
            ) : null}
          </div>
          <button type="button" onClick={onClose} className="rounded-full p-1 text-stone-400" aria-label="Close">
            <X className="size-4" />
          </button>
        </div>

        {error ? (
          <p className="mt-3 inline-flex items-start gap-2 rounded-xl border border-rose-200 bg-rose-50 px-3 py-2.5 text-sm leading-5 text-rose-900">
            <AlertTriangle className="mt-0.5 size-4 shrink-0" />
            {error}
          </p>
        ) : null}

        <form onSubmit={handleSubmit} className="mt-4 space-y-4">
          <BulkField
            label="Quantity on hand"
            enabled={setQuantity}
            onToggle={setSetQuantity}
          >
            <input
              type="number"
              min={0}
              value={quantityValue}
              onChange={(event) => setQuantityValue(event.target.value)}
              disabled={!setQuantity}
              className="h-10 w-full rounded-lg border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100 disabled:bg-stone-50"
            />
          </BulkField>

          <BulkField label="Bin location" enabled={setBinLocation} onToggle={setSetBinLocation}>
            <input
              value={binLocationValue}
              onChange={(event) => setBinLocationValue(event.target.value)}
              disabled={!setBinLocation}
              className="h-10 w-full rounded-lg border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100 disabled:bg-stone-50"
            />
          </BulkField>

          <BulkField label="Category" enabled={setCategory} onToggle={setSetCategory}>
            <input
              value={categoryValue}
              onChange={(event) => setCategoryValue(event.target.value)}
              disabled={!setCategory}
              className="h-10 w-full rounded-lg border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100 disabled:bg-stone-50"
            />
          </BulkField>

          <div className="grid grid-cols-2 gap-3">
            <label className="block">
              <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
                Add tag
              </span>
              <select
                value={addTag}
                onChange={(event) => setAddTag(event.target.value)}
                className="h-10 w-full rounded-lg border border-stone-200 bg-white px-2 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
              >
                <option value={NONE}>Don&apos;t add</option>
                {tags.map((tag) => (
                  <option key={tag} value={tag}>
                    {tag}
                  </option>
                ))}
              </select>
            </label>
            <label className="block">
              <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
                Remove tag
              </span>
              <select
                value={removeTag}
                onChange={(event) => setRemoveTag(event.target.value)}
                className="h-10 w-full rounded-lg border border-stone-200 bg-white px-2 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
              >
                <option value={NONE}>Don&apos;t remove</option>
                {tags.map((tag) => (
                  <option key={tag} value={tag}>
                    {tag}
                  </option>
                ))}
              </select>
            </label>
          </div>
          {tags.length === 0 ? (
            <p className="text-xs text-stone-500">
              No tags yet — add some from Settings to enable bulk tagging.
            </p>
          ) : null}

          <button
            type="submit"
            disabled={submitting || (!fieldsSelected && !tagsSelected)}
            className="w-full rounded-xl bg-stone-900 px-3 py-2.5 text-sm font-semibold text-white disabled:opacity-60"
          >
            {submitting ? "Saving…" : "Apply to selected parts"}
          </button>
        </form>
      </div>
    </div>
  );
}

function BulkField({
  label,
  enabled,
  onToggle,
  children,
}: {
  label: string;
  enabled: boolean;
  onToggle: (value: boolean) => void;
  children: React.ReactNode;
}) {
  return (
    <div>
      <label className="mb-1 flex items-center gap-2">
        <input
          type="checkbox"
          checked={enabled}
          onChange={(event) => onToggle(event.target.checked)}
          className="size-4 rounded border-stone-300 text-amber-700 focus:ring-amber-400"
        />
        <span className="text-xs font-semibold uppercase tracking-wide text-stone-500">
          Set {label.toLowerCase()} to
        </span>
      </label>
      {children}
    </div>
  );
}
