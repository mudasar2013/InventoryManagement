import type { Metadata } from "next";
import { AddPartForm } from "@/components/AddPartForm";

export const metadata: Metadata = {
  title: "Add a part",
};

// A thin wrapper — the writable-sources list AddPartForm needs already
// lives in the client-side InventoryProvider (seeded from
// loadInventory() in app/layout.tsx), so there's nothing server-side
// left to fetch here.
export default function AddPartPage() {
  return <AddPartForm />;
}
