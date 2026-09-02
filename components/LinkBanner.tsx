"use client";

import { Link2 } from "lucide-react";
import { useInventory } from "./InventoryProvider";

export function LinkBanner() {
  const { lastLinkMessage } = useInventory();
  if (!lastLinkMessage) return null;

  return (
    <p
      role="status"
      className="rounded-2xl border border-amber-200 bg-amber-50 px-3.5 py-3 text-sm leading-5 text-amber-950"
    >
      <span className="inline-flex items-start gap-2">
        <Link2 className="mt-0.5 size-4 shrink-0" />
        {lastLinkMessage}
      </span>
    </p>
  );
}
