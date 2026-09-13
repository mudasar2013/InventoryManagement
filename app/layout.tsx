import type { Metadata, Viewport } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import { getServerSession } from "next-auth/next";
import { AlertTriangle } from "lucide-react";
import { BottomNav } from "@/components/BottomNav";
import { InventoryProvider } from "@/components/InventoryProvider";
import { SignOutButton } from "@/components/SignOutButton";
import { authOptions } from "@/lib/auth/options";
import { loadInventory } from "@/lib/getInventory";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: {
    default: "Parts Inventory",
    template: "%s · Parts Inventory",
  },
  description:
    "Look up appliance parts, bin locations, and stock levels for open service jobs.",
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  themeColor: "#f6f3ee",
};

export default async function RootLayout({ children }: LayoutProps<"/">) {
  const session = await getServerSession(authOptions);
  const { parts, jobs, jobParts, warnings, sourceStatuses, tags } = await loadInventory(
    session?.accessToken,
  );

  return (
    <html
      lang="en"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full bg-background font-sans text-foreground">
        <InventoryProvider
          initialParts={parts}
          initialJobs={jobs}
          initialJobParts={jobParts}
          initialSourceStatuses={sourceStatuses}
          initialTags={tags}
        >
          <div className="mx-auto flex min-h-full w-full max-w-lg flex-col px-5 pb-24 pt-7 sm:px-7">
            {session?.user ? (
              <div className="mb-3 flex items-center justify-between gap-2 text-xs text-stone-500">
                <span className="truncate">
                  Signed in as {session.user.name ?? session.user.email}
                </span>
                <SignOutButton />
              </div>
            ) : null}
            {session?.error === "RefreshAccessTokenError" ? (
              <p className="mb-3 inline-flex items-start gap-2 rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-3 text-sm leading-5 text-rose-900">
                <AlertTriangle className="mt-0.5 size-4 shrink-0" />
                Your Microsoft sign-in expired. Sign out and back in to
                restore SharePoint inventory.
              </p>
            ) : null}
            {warnings.map((warning) => (
              <p
                key={warning}
                className="mb-3 inline-flex items-start gap-2 rounded-2xl border border-amber-200 bg-amber-50 px-3.5 py-3 text-sm leading-5 text-amber-950"
              >
                <AlertTriangle className="mt-0.5 size-4 shrink-0" />
                {warning}
              </p>
            ))}
            {children}
          </div>
          <BottomNav />
        </InventoryProvider>
      </body>
    </html>
  );
}
