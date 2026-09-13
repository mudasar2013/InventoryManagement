"use client";

/**
 * A single labeled text/number input, styled to match the rest of the
 * app's forms (see DataSourcesView.tsx, which has its own near-
 * identical local `Field` — this is the shared version for the part
 * edit/add forms, which need the exact same fields in two places:
 * PartDetail.tsx's "Edit this part" form and AddPartForm.tsx).
 */
export function FormField({
  label,
  placeholder,
  value,
  onChange,
  required = true,
  type = "text",
}: {
  label: string;
  placeholder?: string;
  value: string;
  onChange: (value: string) => void;
  required?: boolean;
  type?: "text" | "number";
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-stone-500">
        {label}
        {required ? null : " (optional)"}
      </span>
      <input
        required={required}
        type={type}
        inputMode={type === "number" ? "numeric" : undefined}
        min={type === "number" ? 0 : undefined}
        value={value}
        placeholder={placeholder}
        onChange={(event) => onChange(event.target.value)}
        className="h-11 w-full rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
      />
    </label>
  );
}
