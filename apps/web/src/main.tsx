import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App.tsx";
import "./styles.css";

const root = document.getElementById("root");

if (!root) {
  throw new Error("Application root not found.");
}

// The portfolio demo has no server: before anything renders, requests for /api/v1 are answered
// from a recording. The condition is a build-time constant, so a normal build drops this branch
// together with the shim and the recording.
async function installDemo() {
  const [{ installReplay }, { default: recording }] = await Promise.all([
    import("./demo/replay.ts"),
    import("../demo/fixtures.json?url"),
  ]);
  const response = await fetch(recording);
  if (!response.ok) throw new Error("The demo recording could not be loaded.");
  installReplay(await response.json());
}

async function start(target: HTMLElement) {
  if (import.meta.env.VITE_DEMO === "true") {
    try {
      await installDemo();
    } catch {
      target.textContent = "The demo could not load its recording. Reload the page to try again.";
      return;
    }
  }
  createRoot(target).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}

void start(root);
