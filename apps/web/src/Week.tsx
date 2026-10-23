import type { PrismSession } from "@prism365/contracts";
import type { WorkspaceView } from "./App.tsx";

// The weekly "What changed" page (placeholder until the digest slice lands).
export function ThisWeek(_props: {
  session: PrismSession;
  onNavigate: (view: WorkspaceView) => void;
  onOpenPerson: (userId: string) => void;
}) {
  return <p className="empty-data">This page is being built.</p>;
}
