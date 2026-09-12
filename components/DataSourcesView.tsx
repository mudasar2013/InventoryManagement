import { AlertTriangle, CheckCircle2, CircleSlash, Database } from "lucide-react";
import { AppHeader } from "@/components/AppHeader";
import type { SourceStatus } from "@/lib/getInventory";

/**
 * Visual state for one source's badge. Four states, not just ok/error:
 * a source can be entirely unconfigured (no env vars set at all), or
 * configured but never attempted this request (e.g. SharePoint configured
 * with no Graph token yet), which reads very differently from a source
 * that was attempted and actually failed.
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

export function DataSourcesView({ statuses }: { statuses: SourceStatus[] }) {
  return (
    <main className="space-y-6">
      <AppHeader
        title="Data sources"
        subtitle="Everything the parts catalog reads from, and whether it's connected."
      />

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
                <span
                  className={`inline-flex shrink-0 items-center gap-1 rounded-full px-2.5 py-1 text-xs font-semibold ring-1 ring-inset ${classes}`}
                >
                  <Icon className="size-3.5" />
                  {label}
                </span>
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

      <section className="rounded-2xl border border-dashed border-stone-300 bg-white p-4">
        <h2 className="text-sm font-semibold text-stone-900">Add a data source</h2>
        <p className="mt-1.5 text-sm leading-6 text-stone-600">
          Adding a source today (another SharePoint workbook, or a
          different system entirely) means wiring it into the code under{" "}
          <code className="rounded bg-stone-100 px-1 py-0.5 text-[13px]">
            lib/sources/
          </code>
          , setting its configuration as environment variables, and
          redeploying — the same way the SharePoint source above was set
          up. Adding sources from this page directly, without a redeploy,
          is planned next.
        </p>
      </section>
    </main>
  );
}
