import { PackageX } from "lucide-react";
import Link from "next/link";

export default function NotFound() {
  return (
    <main className="flex flex-1 flex-col items-center justify-center py-20 text-center">
      <PackageX className="size-10 text-stone-400" />
      <h1 className="mt-4 text-xl font-semibold text-stone-900">Page not found</h1>
      <p className="mt-2 max-w-xs text-sm leading-6 text-stone-500">
        That part or job is not in this inventory. Head back to the shelf list.
      </p>
      <Link
        href="/"
        className="mt-6 rounded-full bg-stone-900 px-5 py-2.5 text-sm font-semibold text-white"
      >
        View parts
      </Link>
    </main>
  );
}
