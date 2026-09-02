export function AppHeader({
  title,
  subtitle,
}: {
  title: string;
  subtitle?: string;
}) {
  return (
    <header className="space-y-1">
      <p className="text-xs font-semibold uppercase tracking-[0.16em] text-amber-700">
        Shop floor
      </p>
      <h1 className="text-2xl font-semibold tracking-tight text-stone-900">{title}</h1>
      {subtitle ? <p className="text-sm leading-6 text-stone-600">{subtitle}</p> : null}
    </header>
  );
}
