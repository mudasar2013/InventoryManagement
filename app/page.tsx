import { AppHeader } from "@/components/AppHeader";
import { PartsExplorer } from "@/components/PartsExplorer";

export default function HomePage() {
  return (
    <main className="space-y-6">
      <AppHeader
        title="Parts inventory"
        subtitle="Search any part number without opening a job. Bin, quantity, and stock status stay on the card."
      />
      <PartsExplorer />
    </main>
  );
}
