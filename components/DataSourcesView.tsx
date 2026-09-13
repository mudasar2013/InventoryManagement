"use client";

import {
  AlertTriangle,
  CheckCircle2,
  CircleSlash,
  Database,
  ExternalLink,
  Pencil,
  Plus,
  RefreshCw,
  Trash2,
  X,
} from "lucide-react";
import { useRouter } from "next/navigation";
import { useState, useTransition } from "react";
import { AppHeader } from "@/components/AppHeader";
import type { SourceStatus } from "@/lib/getInventory";
import type { StoredSharePointSource } from "@/lib/sources/sharepoint-source-store";

/**
 * Visual state for one source's badge. Four states, not just ok/error:
 * a source can be entirely unconfigured, or configured but never
 * attempted this request (e.g. SharePoint configured with no Graph
 * token yet), which reads very differently from a source that was
 * attempted and actually failed.
 */
function statusMeta(status: SourceStatus) {
  if (!status.configured) {
    return {
      label: "Not configured",
      classes: "bg-stone-100 text-stone-600 ring-stone-200",
      Icon: CircleSlash,
    };
  }
  if (status.ok === true) {
    return {
      label: "Connected",
      classes: "bg-emerald-50 text-emerald-800 ring-emerald-200",
      Icon: CheckCircle2,
    };
  }
  if (status.ok === false) {
    return {
      label: "Error",
      classes: "bg-rose-50 text-rose-800 ring-rose-200",
      Icon: AlertTriangle,
    };
  }
  return {
    label: "Not connected",
    classes: "bg-amber-50 text-amber-800 ring-amber-200",
    Icon: AlertTriangle,
  };
}

type SourceForm = {
  label: string;
  siteHostname: string;
  sitePath: string;
  filePath: string;
  tableName: string;
  partNumberColumn: string;
  quantityColumn: string;
  descriptionColumn: string;
  binLocationColumn: string;
  idColumn: string;
  categoryColumn: string;
};

const emptyForm: SourceForm = {
  label: "",
  siteHostname: "",
  sitePath: "",
  filePath: "",
  tableName: "",
  partNumberColumn: "",
  quantityColumn: "",
  descriptionColumn: "",
  binLocationColumn: "",
  idColumn: "",
  categoryColumn: "",
};

export function DataSourcesView({
  statuses,
  canAddSharePointSource,
}: {
  statuses: SourceStatus[];
  canAddSharePointSource: boolean;
}) {
  const router = useRouter();
  const [isRefreshing, startRefresh] = useTransition();
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState(emptyForm);
  const [submitting, setSubmitting] = useState(false);
  const [removingId, setRemovingId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Editing an existing source is a two-step affair: fetch its current
  // (untruncated) field values from the collection endpoint — the
  // status list only carries a truncated `detail` string, not the raw
  // siteHostname/sitePath/filePath/tableName a form needs — then show
  // those in the same field layout as the add form.
  const [loadingEditId, setLoadingEditId] = useState<string | null>(null);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState<SourceForm>(emptyForm);
  const [editSubmitting, setEditSubmitting] = useState(false);

  function handleRefresh() {
    setError(null);
    startRefresh(() => {
      router.refresh();
    });
  }

  async function handleAdd(event: React.FormEvent) {
    event.preventDefault();
    setSubmitting(true);
    setError(null);
    try {
      const response = await fetch("/api/sources/sharepoint", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(form),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to add source.");
      }
      setForm(emptyForm);
      setOpen(false);
      router.refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to add source.");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleRemove(id: string) {
    if (editingId === id) {
      setEditingId(null);
    }
    setRemovingId(id);
    setError(null);
    try {
      const response = await fetch(`/api/sources/sharepoint/${id}`, {
        method: "DELETE",
      });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error ?? "Failed to remove source.");
      }
      router.refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to remove source.");
    } finally {
      setRemovingId(null);
    }
  }

  async function handleStartEdit(status: SourceStatus) {
    setError(null);
    setOpen(false);
    setLoadingEditId(status.id);
    try {
      const response = await fetch("/api/sources/sharepoint");
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to load source details.");
      }
      const sources = (payload.sources ?? []) as StoredSharePointSource[];
      const match = sources.find((source) => source.id === status.id);
      if (!match) {
        throw new Error("Couldn't find that source's details — it may have just been removed.");
      }
      setEditForm({
        label: match.label,
        siteHostname: match.siteHostname,
        sitePath: match.sitePath,
        filePath: match.filePath,
        tableName: match.tableName,
        // Sources added before per-source column mapping existed have
        // none of these set — prefill with the legacy default names so
        // editing one doesn't present blank required fields.
        partNumberColumn: match.partNumberColumn ?? "PartNumber",
        quantityColumn: match.quantityColumn ?? "QtyOnHand",
        descriptionColumn: match.descriptionColumn ?? "",
        binLocationColumn: match.binLocationColumn ?? "",
        idColumn: match.idColumn ?? "",
        categoryColumn: match.categoryColumn ?? "",
      });
      setEditingId(status.id);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load source details.");
    } finally {
      setLoadingEditId(null);
    }
  }

  function handleCancelEdit() {
    setEditingId(null);
    setError(null);
  }

  async function handleSaveEdit(event: React.FormEvent) {
    event.preventDefault();
    if (!editingId) {
      return;
    }
    setEditSubmitting(true);
    setError(null);
    try {
      const response = await fetch(`/api/sources/sharepoint/${editingId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(editForm),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to update source.");
      }
      setEditingId(null);
      router.refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to update source.");
    } finally {
      setEditSubmitting(false);
    }
  }

  return (
    <main className="space-y-6">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <AppHeader
            title="Data sources"
            subtitle="Everything the parts catalog reads from, and whether it's connected."
          />
        </div>
        <button
          type="button"
          onClick={handleRefresh}
          disabled={isRefreshing}
          className="mt-1 inline-flex shrink-0 items-center gap-1.5 rounded-full border border-stone-200 bg-white px-3 py-1.5 text-xs font-semibold text-stone-700 disabled:opacity-60"
        >
          <RefreshCw className={`size-3.5 ${isRefreshing ? "animate-spin" : ""}`} />
          {isRefreshing ? "Refreshing…" : "Refresh"}
        </button>
      </div>

      {error ? (
        <p className="inline-flex items-start gap-2 rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-3 text-sm leading-5 text-rose-900">
          <AlertTriangle className="mt-0.5 size-4 shrink-0" />
          {error}
        </p>
      ) : null}

      <ul className="space-y-3">
        {statuses.map((status) => {
          const { label, classes, Icon } = statusMeta(status);
          const isEditing = editingId === status.id;
          return (
            <li
              key={status.id}
              className="rounded-2xl border border-stone-200 bg-white p-4 shadow-sm"
            >
              <div className="flex items-start justify-between gap-3">
                <div className="flex min-w-0 items-center gap-3">
                  <div className="flex size-11 shrink-0 items-center justify-center rounded-xl bg-amber-50 text-amber-700">
                    <Database className="size-5" />
                  </div>
                  <div className="min-w-0">
                    <p className="text-sm font-semibold text-stone-900">{status.label}</p>
                    {status.detail ? (
                      <p className="mt-0.5 truncate text-xs text-stone-500">
                        {status.detail}
                      </p>
                    ) : null}
                  </div>
                </div>
                <div className="flex shrink-0 items-center gap-2">
                  <span
                    className={`inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-semibold ring-1 ring-inset ${classes}`}
                  >
                    <Icon className="size-3.5" />
                    {label}
                  </span>
                  {status.removable ? (
                    <>
                      <button
                        type="button"
                        onClick={() => handleStartEdit(status)}
                        disabled={loadingEditId === status.id}
                        aria-label={`Edit ${status.label}`}
                        className="rounded-full p-1.5 text-stone-400 hover:bg-amber-50 hover:text-amber-700 disabled:opacity-50"
                      >
                        <Pencil className="size-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => handleRemove(status.id)}
                        disabled={removingId === status.id}
                        aria-label={`Remove ${status.label}`}
                        className="rounded-full p-1.5 text-stone-400 hover:bg-rose-50 hover:text-rose-600 disabled:opacity-50"
                      >
                        <Trash2 className="size-3.5" />
                      </button>
                    </>
                  ) : null}
                </div>
              </div>

              {status.fileUrl ? (
                <a
                  href={status.fileUrl}
                  target="_blank"
                  rel="noreferrer"
                  className="mt-3 inline-flex items-center gap-1.5 text-xs font-semibold text-amber-700 hover:underline"
                >
                  <ExternalLink className="size-3.5" />
                  Open source file
                </a>
              ) : null}

              {typeof status.partCount === "number" ? (
                <p className="mt-3 text-xs font-medium text-stone-500">
                  {status.partCount} part{status.partCount === 1 ? "" : "s"} reported
                </p>
              ) : null}

              {status.note ? (
                <p className="mt-2 text-xs leading-5 text-stone-600">{status.note}</p>
              ) : null}

              {status.debugDetail ? (
                <details className="mt-2 text-xs">
                  <summary className="cursor-pointer select-none font-medium text-stone-500 hover:text-stone-700">
                    Show details
                  </summary>
                  <pre className="mt-2 max-h-56 overflow-auto whitespace-pre-wrap break-words rounded-lg bg-stone-50 p-2.5 text-[11px] leading-4 text-stone-700">
                    {status.debugDetail}
                  </pre>
                </details>
              ) : null}

              {isEditing ? (
                <form
                  onSubmit={handleSaveEdit}
                  className="mt-4 space-y-3 border-t border-stone-100 pt-4"
                >
                  <div className="flex items-center justify-between gap-2">
                    <p className="text-sm font-semibold text-stone-900">
                      Edit this source
                    </p>
                    <button
                      type="button"
                      onClick={handleCancelEdit}
                      className="rounded-full p-1 text-stone-400"
                      aria-label="Cancel edit"
                    >
                      <X className="size-4" />
                    </button>
                  </div>

                  <Field
                    label="Name"
                    placeholder="e.g. Downtown shop inventory"
                    value={editForm.label}
                    onChange={(value) => setEditForm((f) => ({ ...f, label: value }))}
                  />
                  <Field
                    label="Site hostname"
                    placeholder="contoso.sharepoint.com"
                    value={editForm.siteHostname}
                    onChange={(value) =>
                      setEditForm((f) => ({ ...f, siteHostname: value }))
                    }
                  />
                  <Field
                    label="Site path"
                    placeholder="/sites/ServiceOps"
                    value={editForm.sitePath}
                    onChange={(value) => setEditForm((f) => ({ ...f, sitePath: value }))}
                  />
                  <Field
                    label="File path"
                    placeholder="Shared Documents/Inventory.xlsx"
                    value={editForm.filePath}
                    onChange={(value) => setEditForm((f) => ({ ...f, filePath: value }))}
                  />
                  <Field
                    label="Table or sheet name"
                    placeholder="Parts (or a worksheet tab name, e.g. Sheet1)"
                    value={editForm.tableName}
                    onChange={(value) => setEditForm((f) => ({ ...f, tableName: value }))}
                  />

                  <p className="pt-1 text-xs leading-5 text-stone-500">
                    This sheet&apos;s own column headers (row 1) — every source can use
                    different names. Check &quot;Show details&quot; above for the exact
                    headers this sheet actually has if a fetch failed on this.
                  </p>
                  <Field
                    label="Part number column"
                    placeholder="PartNumber"
                    value={editForm.partNumberColumn}
                    onChange={(value) =>
                      setEditForm((f) => ({ ...f, partNumberColumn: value }))
                    }
                  />
                  <Field
                    label="Quantity column"
                    placeholder="QtyOnHand"
                    value={editForm.quantityColumn}
                    onChange={(value) =>
                      setEditForm((f) => ({ ...f, quantityColumn: value }))
                    }
                  />
                  <Field
                    label="Description column"
                    placeholder="Description"
                    value={editForm.descriptionColumn}
                    onChange={(value) =>
                      setEditForm((f) => ({ ...f, descriptionColumn: value }))
                    }
                    required={false}
                  />
                  <Field
                    label="Location / bin column"
                    placeholder="BinLocation"
                    value={editForm.binLocationColumn}
                    onChange={(value) =>
                      setEditForm((f) => ({ ...f, binLocationColumn: value }))
                    }
                    required={false}
                  />
                  <Field
                    label="Category column"
                    placeholder="Category"
                    value={editForm.categoryColumn}
                    onChange={(value) =>
                      setEditForm((f) => ({ ...f, categoryColumn: value }))
                    }
                    required={false}
                  />
                  <Field
                    label="Id column"
                    placeholder="Leave blank to generate one from the part number"
                    value={editForm.idColumn}
                    onChange={(value) => setEditForm((f) => ({ ...f, idColumn: value }))}
                    required={false}
                  />

                  <div className="flex gap-2">
                    <button
                      type="button"
                      onClick={handleCancelEdit}
                      className="flex-1 rounded-xl border border-stone-200 px-3 py-2.5 text-sm font-semibold text-stone-700"
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      disabled={editSubmitting}
                      className="flex-1 rounded-xl bg-stone-900 px-3 py-2.5 text-sm font-semibold text-white disabled:opacity-60"
                    >
                      {editSubmitting ? "Saving…" : "Save & refresh"}
                    </button>
                  </div>
                </form>
              ) : null}
            </li>
          );
        })}
      </ul>

      {canAddSharePointSource ? (
        open ? (
          <form
            onSubmit={handleAdd}
            className="space-y-3 rounded-2xl border border-stone-200 bg-white p-4"
          >
            <div className="flex items-center justify-between gap-2">
              <p className="text-sm font-semibold text-stone-900">
                Add a SharePoint workbook
              </p>
              <button
                type="button"
                onClick={() => setOpen(false)}
                className="rounded-full p-1 text-stone-400"
                aria-label="Cancel"
              >
                <X className="size-4" />
              </button>
            </div>

            <Field
              label="Name"
              placeholder="e.g. Downtown shop inventory"
              value={form.label}
              onChange={(value) => setForm((f) => ({ ...f, label: value }))}
            />
            <Field
              label="Site hostname"
              placeholder="contoso.sharepoint.com"
              value={form.siteHostname}
              onChange={(value) => setForm((f) => ({ ...f, siteHostname: value }))}
            />
            <Field
              label="Site path"
              placeholder="/sites/ServiceOps"
              value={form.sitePath}
              onChange={(value) => setForm((f) => ({ ...f, sitePath: value }))}
            />
            <Field
              label="File path"
              placeholder="Shared Documents/Inventory.xlsx"
              value={form.filePath}
              onChange={(value) => setForm((f) => ({ ...f, filePath: value }))}
            />
            <Field
              label="Table or sheet name"
              placeholder="Parts (or a worksheet tab name, e.g. Sheet1)"
              value={form.tableName}
              onChange={(value) => setForm((f) => ({ ...f, tableName: value }))}
            />

            <p className="pt-1 text-xs leading-5 text-stone-500">
              This sheet&apos;s own column headers (row 1) — every source can use
              different names, so this isn&apos;t shared across sources.
            </p>
            <Field
              label="Part number column"
              placeholder="PartNumber"
              value={form.partNumberColumn}
              onChange={(value) => setForm((f) => ({ ...f, partNumberColumn: value }))}
            />
            <Field
              label="Quantity column"
              placeholder="QtyOnHand"
              value={form.quantityColumn}
              onChange={(value) => setForm((f) => ({ ...f, quantityColumn: value }))}
            />
            <Field
              label="Description column"
              placeholder="Description"
              value={form.descriptionColumn}
              onChange={(value) => setForm((f) => ({ ...f, descriptionColumn: value }))}
              required={false}
            />
            <Field
              label="Location / bin column"
              placeholder="BinLocation"
              value={form.binLocationColumn}
              onChange={(value) => setForm((f) => ({ ...f, binLocationColumn: value }))}
              required={false}
            />
            <Field
              label="Category column"
              placeholder="Category"
              value={form.categoryColumn}
              onChange={(value) => setForm((f) => ({ ...f, categoryColumn: value }))}
              required={false}
            />
            <Field
              label="Id column"
              placeholder="Leave blank to generate one from the part number"
              value={form.idColumn}
              onChange={(value) => setForm((f) => ({ ...f, idColumn: value }))}
              required={false}
            />

            <button
              type="submit"
              disabled={submitting}
              className="w-full rounded-xl bg-stone-900 px-3 py-2.5 text-sm font-semibold text-white disabled:opacity-60"
            >
              {submitting ? "Adding…" : "Add source"}
            </button>
          </form>
        ) : (
          <button
            type="button"
            onClick={() => {
              setEditingId(null);
              setOpen(true);
            }}
            className="inline-flex w-full items-center justify-center gap-1.5 rounded-xl border border-dashed border-stone-300 px-3 py-2.5 text-sm font-semibold text-stone-700"
          >
            <Plus className="size-4" />
            Add a SharePoint workbook
          </button>
        )
      ) : (
        <section className="rounded-2xl border border-dashed border-stone-300 bg-white p-4">
          <h2 className="text-sm font-semibold text-stone-900">Add a data source</h2>
          <p className="mt-1.5 text-sm leading-6 text-stone-600">
            Adding a SharePoint workbook from this page needs somewhere to
            save it — this project doesn&apos;t have a Redis store yet. Add
            a Redis integration in the Vercel dashboard (Project → Storage
            → Marketplace Database Providers), then this form will appear
            here.
          </p>
        </section>
      )}
    </main>
  );
}

function Field({
  label,
  placeholder,
  value,
  onChange,
  required = true,
}: {
  label: string;
  placeholder: string;
  value: string;
  onChange: (value: string) => void;
  required?: boolean;
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
        {label}
        {required ? null : " (optional)"}
      </span>
      <input
        required={required}
        value={value}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
        className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
      />
    </label>
  );
}
