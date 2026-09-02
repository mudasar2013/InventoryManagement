export default function Loading() {
  return (
    <main className="space-y-6" aria-busy="true" aria-label="Loading inventory">
      <div className="space-y-2">
        <div className="h-3 w-20 animate-pulse rounded bg-stone-200" />
        <div className="h-8 w-48 animate-pulse rounded-lg bg-stone-200" />
        <div className="h-4 w-full animate-pulse rounded bg-stone-200" />
      </div>
      <div className="h-12 animate-pulse rounded-2xl bg-stone-200" />
      <div className="space-y-3">
        <div className="h-24 animate-pulse rounded-2xl bg-stone-200" />
        <div className="h-24 animate-pulse rounded-2xl bg-stone-200" />
        <div className="h-24 animate-pulse rounded-2xl bg-stone-200" />
      </div>
    </main>
  );
}
