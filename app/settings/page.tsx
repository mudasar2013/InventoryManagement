import type { Metadata } from "next";
import { getServerSession } from "next-auth/next";
import { SettingsView } from "@/components/SettingsView";
import { authOptions } from "@/lib/auth/options";
import { loadInventory } from "@/lib/getInventory";

export const metadata: Metadata = {
  title: "Settings",
};

export default async function SettingsPage() {
  const session = await getServerSession(authOptions);
  const { tags, canManageTags } = await loadInventory(session?.accessToken);

  return <SettingsView tags={tags} canManageTags={canManageTags} />;
}
