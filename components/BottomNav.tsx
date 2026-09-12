"use client";

import { ClipboardList, Database, Package } from "lucide-react";
import Link from "next/link";
import { usePathname } from "next/navigation";

const tabs = [
  { href: "/", label: "Parts", icon: Package, match: (path: string) => path === "/" || path.startsWith("/parts") },
  { href: "/jobs", label: "Jobs", icon: ClipboardList, match: (path: string) => path.startsWith("/jobs") },
  { href: "/sources", label: "Sources", icon: Database, match: (path: string) => path.startsWith("/sources") },
];

export function BottomNav() {
  const pathname = usePathname();

  return (
    <nav
      aria-label="Primary"
      className="fixed inset-x-0 bottom-0 z-20 border-t border-stone-200 bg-white/95 backdrop-blur-md"
    >
      <ul className="mx-auto flex h-16 max-w-lg items-stretch">
        {tabs.map((tab) => {
          const active = tab.match(pathname);
          const Icon = tab.icon;
          return (
            <li key={tab.href} className="flex-1">
              <Link
                href={tab.href}
                className={`flex h-full flex-col items-center justify-center gap-1 text-xs font-semibold ${
                  active ? "text-amber-700" : "text-stone-500"
                }`}
              >
                <Icon className="size-5" strokeWidth={active ? 2.4 : 2} />
                {tab.label}
              </Link>
            </li>
          );
        })}
      </ul>
    </nav>
  );
}
