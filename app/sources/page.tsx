import type { Metadata } from "next";
import { getServerSession } from "next-auth/next";
import { DataSourcesView } from "@/components/DataSourcesView";
import { authOptions } from "@/lib/auth/options";
import { loadInventory } from "@/lib/getInventory";

export const metadata: Metadata = {
  title: "Data sources",
};

export default async function SourcesPage() {
  const session = await getServerSession(authOptions);
  const { sourceStatuses, canAddSharePointSource } = await loadInventory(
    session?.accessToken,
  );

  return (
    <DataSourcesView
      statuses={sourceStatuses}
      canAddSharePointSource={canAddSharePointSource}
    />
  );
}
