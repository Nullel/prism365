import { type MouseEvent, useCallback, useEffect, useRef, useState } from "react";
import { CheckIcon, CopyIcon } from "./icons.tsx";

// One-click copy for the addresses in audit tables and the people list. Every copy action in the
// app goes through copyText, so a browser that blocks the clipboard API (a page served over plain
// http, an embedded view, a denied permission) still copies when the browser allows it at all.

const FAILED_MILLISECONDS = 4_000;

// Copies through a hidden, read-only textarea. The old route, kept for browsers and pages where
// navigator.clipboard is missing or refuses.
function copyThroughTextarea(text: string): boolean {
  if (typeof document === "undefined") return false;
  const area = document.createElement("textarea");
  // Focus returns to where the administrator was, so a keyboard user keeps their place.
  const previous = document.activeElement as { focus?: (options?: FocusOptions) => void } | null;
  try {
    area.value = text;
    area.setAttribute("readonly", "");
    area.setAttribute("aria-hidden", "true");
    area.style.position = "fixed";
    area.style.top = "0";
    area.style.left = "0";
    area.style.opacity = "0";
    document.body.appendChild(area);
    area.select();
    area.setSelectionRange(0, text.length);
    return document.execCommand("copy");
  } catch {
    return false;
  } finally {
    area.remove();
    try {
      previous?.focus?.({ preventScroll: true });
    } catch {
      // Nothing to put the focus back on.
    }
  }
}

// Puts the text on the clipboard. Resolves to whether it worked; never throws.
export async function copyText(text: string): Promise<boolean> {
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // Refused or unavailable: try the textarea route below.
  }
  try {
    return copyThroughTextarea(text);
  } catch {
    return false;
  }
}

export type CopyStatus = "idle" | "copied" | "failed";

// Runs a copy and reports "copied" for resetMilliseconds (longer for "failed", so the way out can
// be read), then goes back to "idle". A new copy restarts the clock; leaving the page stops it.
export function useCopy(resetMilliseconds: number) {
  const [status, setStatus] = useState<CopyStatus>("idle");
  const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const mounted = useRef(false);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      clearTimeout(timer.current);
    };
  }, []);

  const copy = useCallback(
    async (text: string): Promise<boolean> => {
      const copied = await copyText(text);
      if (!mounted.current) return copied;
      clearTimeout(timer.current);
      setStatus(copied ? "copied" : "failed");
      timer.current = setTimeout(
        () => setStatus("idle"),
        copied ? resetMilliseconds : Math.max(resetMilliseconds, FAILED_MILLISECONDS),
      );
      return copied;
    },
    [resetMilliseconds],
  );

  return { status, copy };
}

const FAILED_COPY_HINT = "Copy failed. Select the text and press Ctrl+C or Cmd+C.";

// A small icon button that copies `text`. It stops the click here, so it never opens the row or
// link it sits in. The polite status line tells a screen reader what the icon shows.
export function CopyButton({ text, label }: { text: string; label?: string }) {
  const { status, copy } = useCopy(1_500);
  const name = label ?? `Copy ${text}`;
  const onClick = (event: MouseEvent<HTMLButtonElement>) => {
    event.stopPropagation();
    void copy(text);
  };
  return (
    <>
      <button
        className={`copy-button${status === "idle" ? "" : ` is-${status}`}`}
        type="button"
        aria-label={name}
        title={status === "failed" ? FAILED_COPY_HINT : name}
        onClick={onClick}
      >
        {status === "copied" ? <CheckIcon /> : <CopyIcon />}
      </button>
      <span className="visually-hidden" aria-live="polite">
        {status === "copied" ? "Copied" : status === "failed" ? "Copy failed" : ""}
      </span>
    </>
  );
}
