import type { PrismSession } from "@prism365/contracts";

// The directory change feed (placeholder until the directory changes slice lands).
export function DirectoryChanges(_props: {
  session: PrismSession;
  onOpenPerson: (userId: string) => void;
}) {
  return <p className="empty-data">This page is being built.</p>;
}
