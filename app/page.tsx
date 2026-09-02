import { AppHeader } from "@/components/AppHeader";
import { PartsExplorer } from "@/components/PartsExplorer";
import { countByStatus } from "@/lib/inventory";
import { parts } from "@/lib/mockData";

export default function HomePage() {
  return (
    <main className="space-y-6">
      <AppHeader
        title="Parts inventory"
        subtitle="Find a part number, pull it from the bin, and see what is still on the shelf."
      />
      <PartsExplorer parts={parts} counts={countByStatus()} />
    </main>
  );
}
