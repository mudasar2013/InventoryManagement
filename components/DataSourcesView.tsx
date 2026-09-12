"use client";

import {
  AlertTriangle,
  CheckCircle2,
  CircleSlash,
  Database,
  Plus,
  Trash2,
  X,
} from "lucide-react";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { AppHeader } from "@/components/AppHeader";
import type { SourceStatus } from "@/lib/getInventory";

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

const emptyForm = {
  label: "",
  siteHostname: "",
  sitePath: "",
  filePath: "",
  tableName: "",
};

export function DataSourcesView({
  statuses,
  canAddSharePointSource,
}: {
  statuses: SourceStatus[];
  canAddSharePointSource: boolean;
}) {
  const router = useRouter();
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState(emptyForm);
  const [submitting, setSubmitting] = useState(false);
  const [removingId, setRemovingId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

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

  return (
    <main className="space-y-6">
      <AppHeader
        title="Data sources"
        subtitle="Everything the parts catalog reads from, and whether it's connected."
      />

      {error ? (
        <p className="inline-flex items-start gap-2 rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-3 text-sm leading-5 text-rose-900">
          <AlertTriangle className="mt-0.5 size-4 shrink-0" />
          {error}
        </p>
      ) : null}

      <ul className="space-y-3">
        {statuses.map((status) => {
          const { label, classes, Icon } = statusMeta(status);
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
                    <button
                      type="button"
                      onClick={() => handleRemove(status.id)}
                      disabled={removingId === status.id}
                      aria-label={`Remove ${status.label}`}
                      className="rounded-full p-1.5 text-stone-400 hover:bg-rose-50 hover:text-rose-600 disabled:opacity-50"
                    >
                      <Trash2 className="size-3.5" />
                    </button>
                  ) : null}
                </div>
              </div>

              {typeof status.partCount === "number" ? (
                <p className="mt-3 text-xs font-medium text-stone-500">
                  {status.partCount} part{status.partCount === 1 ? "" : "s"} reported
                </p>
              ) : null}

              {status.note ? (
                <p className="mt-2 text-xs leading-5 text-stone-600">{status.note}</p>
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
              label="Table name"
              placeholder="Parts"
              value={form.tableName}
              onChange={(value) => setForm((f) => ({ ...f, tableName: value }))}
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
            onClick={() => setOpen(true)}
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
}: {
  label: string;
  placeholder: string;
  value: string;
  onChange: (value: string) => void;
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
        {label}
      </span>
      <input
        required
        value={value}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
        className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
      />
    </label>
  );
}
