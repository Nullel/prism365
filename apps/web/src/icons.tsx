import type { SVGProps } from "react";

type IconProps = SVGProps<SVGSVGElement>;

const common = {
  fill: "none",
  stroke: "currentColor",
  strokeLinecap: "round" as const,
  strokeLinejoin: "round" as const,
  strokeWidth: 1.7,
  viewBox: "0 0 24 24",
};

export function GridIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <rect x="3" y="3" width="7" height="7" rx="2" />
      <rect x="14" y="3" width="7" height="7" rx="2" />
      <rect x="3" y="14" width="7" height="7" rx="2" />
      <rect x="14" y="14" width="7" height="7" rx="2" />
    </svg>
  );
}

export function ShieldIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M12 3 20 6v5c0 5.1-3.3 8.6-8 10-4.7-1.4-8-4.9-8-10V6l8-3Z" />
      <path d="m8.8 12 2.1 2.1 4.6-4.7" />
    </svg>
  );
}

export function UsersIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M16 20v-1.5c0-2-1.8-3.5-4-3.5H6c-2.2 0-4 1.5-4 3.5V20" />
      <circle cx="9" cy="8" r="3" />
      <path d="M16 5.2a3 3 0 0 1 0 5.6M18 15c2.2 0 4 1.5 4 3.5V20" />
    </svg>
  );
}

export function ReportIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M5 3h10l4 4v14H5V3Z" />
      <path d="M15 3v5h4M8 16v-3M12 16V9M16 16v-5" />
    </svg>
  );
}

export function ActivityIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M3 12h4l2-6 4 12 2-6h6" />
    </svg>
  );
}

export function SettingsIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <circle cx="12" cy="12" r="3" />
      <path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1-2.8 2.8-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.6v.2h-4V21a1.7 1.7 0 0 0-1-1.6 1.7 1.7 0 0 0-1.9.3l-.1.1L4.2 17l.1-.1a1.7 1.7 0 0 0 .3-1.9A1.7 1.7 0 0 0 3 14H2.8v-4H3a1.7 1.7 0 0 0 1.6-1 1.7 1.7 0 0 0-.3-1.9L4.2 7 7 4.2l.1.1A1.7 1.7 0 0 0 9 4.6a1.7 1.7 0 0 0 1-1.6v-.2h4V3a1.7 1.7 0 0 0 1 1.6 1.7 1.7 0 0 0 1.9-.3l.1-.1L19.8 7l-.1.1a1.7 1.7 0 0 0-.3 1.9 1.7 1.7 0 0 0 1.6 1h.2v4H21a1.7 1.7 0 0 0-1.6 1Z" />
    </svg>
  );
}

export function SearchIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <circle cx="11" cy="11" r="7" />
      <path d="m20 20-4-4" />
    </svg>
  );
}

export function SparkleIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="m12 3 .8 3.2A4 4 0 0 0 15.7 9l3.3.9-3.3.9a4 4 0 0 0-2.9 2.9L12 17l-.8-3.3a4 4 0 0 0-2.9-2.9L5 10l3.3-.8a4 4 0 0 0 2.9-3L12 3Z" />
      <path d="m19 16 .4 1.3a2 2 0 0 0 1.3 1.3L22 19l-1.3.4a2 2 0 0 0-1.3 1.3L19 22l-.4-1.3a2 2 0 0 0-1.3-1.3L16 19l1.3-.4a2 2 0 0 0 1.3-1.3L19 16Z" />
    </svg>
  );
}

export function ChevronIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="m9 18 6-6-6-6" />
    </svg>
  );
}

export function ArrowUpRightIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M7 17 17 7M8 7h9v9" />
    </svg>
  );
}

export function LockIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <rect x="4" y="10" width="16" height="11" rx="3" />
      <path d="M8 10V7a4 4 0 0 1 8 0v3" />
    </svg>
  );
}

export function BellIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M6 9a6 6 0 1 1 12 0c0 5 2 6.5 2 6.5H4S6 14 6 9" />
      <path d="M10 19a2 2 0 0 0 4 0" />
    </svg>
  );
}

export function ChecklistIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M9 4h6a1 1 0 0 1 1 1v1H8V5a1 1 0 0 1 1-1Z" />
      <path d="M8 5H6a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2" />
      <path d="m8 12 1.5 1.5L12 11M8 17l1.5 1.5L12 16M14.5 12.5H16M14.5 17.5H16" />
    </svg>
  );
}

export function DownloadIcon(props: IconProps) {
  return (
    <svg aria-hidden="true" {...common} {...props}>
      <path d="M12 4v11M7.5 10.5 12 15l4.5-4.5M5 19h14" />
    </svg>
  );
}
