import { ShieldIcon } from "../icons.tsx";

// Shown above every page of the portfolio demo build, and nowhere else. It reuses the notice the
// older synthetic mode uses, so it needs no styles of its own.
export function DemoBanner() {
  return (
    <section className="demo-banner" aria-label="About this demo">
      <div className="demo-banner__icon">
        <ShieldIcon />
      </div>
      <div>
        <span>
          Demo: Larkspur Analytical is an invented company. Every screen is real Prism365 output,
          recorded from the server running against a synthetic Microsoft tenant.
        </span>
      </div>
    </section>
  );
}
