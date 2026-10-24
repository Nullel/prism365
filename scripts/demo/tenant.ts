import { createHash } from "node:crypto";

// A synthetic Microsoft Graph for the portfolio demo.
//
// Larkspur Analytical is an invented analytical laboratory. Nothing here comes from a real
// tenant: every person, address, device and incident is made up, every domain is under the
// reserved .example suffix (RFC 2606) and every IP address sits in a documentation range
// (RFC 5737: 192.0.2.0/24, 198.51.100.0/24, 203.0.113.0/24; RFC 3849: 2001:db8::/32).
//
// createSyntheticGraph() answers exactly the Microsoft Graph v1.0 requests Prism365 makes, so the
// real server code (collectors, audits, directory reads) runs unchanged against it. The data is
// deterministic: people, devices and stories come from a seeded generator, and only the
// timestamps move with `now`.

const MINUTE = 60_000;
const HOUR = 3_600_000;
const DAY = 86_400_000;

export const ORGANISATION = "Larkspur Analytical";
export const DOMAIN = "larkspur.example";
// The research subsidiary signs in on its own domain in the same tenant, so the demo shows an
// audit scope with more than one domain to choose from.
export const RESEARCH_DOMAIN = "larkspur-research.example";
export const SYNTHETIC_TENANT_ID = "6f1c2d8e-4b7a-4e35-9a0c-3d5b7e91c204";
export const SYNTHETIC_ACTOR_ID = guid("user:admin");

// Hosts a synthetic file may mention besides *.example: Microsoft's documented admin center and
// documentation links, which the message and incident text point at.
export const DOCUMENTED_HOSTS = ["admin.microsoft.com", "learn.microsoft.com"] as const;

// ---------------------------------------------------------------------------------------------
// Deterministic helpers
// ---------------------------------------------------------------------------------------------

// A stable GUID for a label, so ids never change between recordings.
export function guid(label: string): string {
  const bytes = Buffer.from(
    createHash("sha256").update(`larkspur:${label}`).digest().subarray(0, 16),
  );
  bytes[6] = ((bytes[6] ?? 0) & 0x0f) | 0x40;
  bytes[8] = ((bytes[8] ?? 0) & 0x3f) | 0x80;
  const hex = bytes.toString("hex");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

function seeded(label: string): () => number {
  let state = createHash("sha256").update(`seed:${label}`).digest().readUInt32LE(0);
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296;
  };
}

type Rand = () => number;
const between = (rand: Rand, low: number, high: number) =>
  low + Math.floor(rand() * (high - low + 1));
const pick = <T>(rand: Rand, items: readonly T[]): T =>
  items[Math.floor(rand() * items.length)] as T;
const chance = (rand: Rand, probability: number) => rand() < probability;
const startOfDay = (ms: number) => ms - (((ms % DAY) + DAY) % DAY);
const iso = (ms: number) => new Date(ms).toISOString();

// ---------------------------------------------------------------------------------------------
// Sites, apps, devices
// ---------------------------------------------------------------------------------------------

type SiteId = "madison" | "minneapolis" | "denver" | "raleigh" | "portland" | "austin";

interface Place {
  city: string;
  state: string;
  country: string;
  latitude: number;
  longitude: number;
}

interface Site extends Place {
  ips: readonly string[];
  // Offset from UTC in hours at the time of the data (US daylight time).
  utcOffset: number;
  office: string;
}

const sites: Record<SiteId, Site> = {
  madison: {
    city: "Madison",
    state: "Wisconsin",
    country: "US",
    latitude: 43.0731,
    longitude: -89.4012,
    ips: ["192.0.2.14", "192.0.2.15"],
    utcOffset: -5,
    office: "Madison laboratory",
  },
  minneapolis: {
    city: "Minneapolis",
    state: "Minnesota",
    country: "US",
    latitude: 44.9778,
    longitude: -93.265,
    ips: ["198.51.100.22", "198.51.100.23"],
    utcOffset: -5,
    office: "Minneapolis office",
  },
  denver: {
    city: "Denver",
    state: "Colorado",
    country: "US",
    latitude: 39.7392,
    longitude: -104.9903,
    ips: ["203.0.113.41"],
    utcOffset: -6,
    office: "Remote",
  },
  raleigh: {
    city: "Raleigh",
    state: "North Carolina",
    country: "US",
    latitude: 35.7796,
    longitude: -78.6382,
    ips: ["203.0.113.52"],
    utcOffset: -4,
    office: "Remote",
  },
  portland: {
    city: "Portland",
    state: "Oregon",
    country: "US",
    latitude: 45.5152,
    longitude: -122.6784,
    ips: ["203.0.113.61"],
    utcOffset: -7,
    office: "Remote",
  },
  austin: {
    city: "Austin",
    state: "Texas",
    country: "US",
    latitude: 30.2672,
    longitude: -97.7431,
    ips: ["203.0.113.71"],
    utcOffset: -5,
    office: "Remote",
  },
};

// Places and addresses used by the planted stories.
const bucharest: Place = {
  city: "Bucharest",
  state: "Bucuresti",
  country: "RO",
  latitude: 44.4268,
  longitude: 26.1025,
};
// A cloud host in Virginia, so the failed prompts do not add a country of their own to the
// sign-in summary: Microsoft counts failed MFA prompts under neither successes nor failures.
const ashburn: Place = {
  city: "Ashburn",
  state: "Virginia",
  country: "US",
  latitude: 39.0438,
  longitude: -77.4874,
};
const frankfurt: Place = {
  city: "Frankfurt am Main",
  state: "Hessen",
  country: "DE",
  latitude: 50.1109,
  longitude: 8.6821,
};
const TRAVEL_IP = "198.51.100.231";
const MFA_ATTACKER_IP = "203.0.113.77";
const SPRAY_IP = "2001:db8:4f1a::9c";
const WENDELL_HOME_IP = "203.0.113.18";
// The address of the run of failed sign-ins before a success, for the recorder to look for.
export const HOME_FAILURE_IP = WENDELL_HOME_IP;

interface AppDefinition {
  name: string;
  resource: string;
}

const apps = {
  exchange: { name: "Office 365 Exchange Online", resource: "Office 365 Exchange Online" },
  teams: { name: "Microsoft Teams", resource: "Microsoft Teams Services" },
  sharepoint: { name: "Office 365 SharePoint Online", resource: "Office 365 SharePoint Online" },
  portal: { name: "Microsoft Office 365 Portal", resource: "Windows Azure Active Directory" },
  onedrive: { name: "OneDrive SyncEngine", resource: "Office 365 SharePoint Online" },
  azure: { name: "Azure Portal", resource: "Windows Azure Service Management API" },
  powerbi: { name: "Power BI Service", resource: "Power BI Service" },
} satisfies Record<string, AppDefinition>;
type AppKey = keyof typeof apps;

interface Device {
  id: string;
  os: string;
  browser: string;
  managed: boolean;
  compliant: boolean | null;
  trustType: string | null;
}

// ---------------------------------------------------------------------------------------------
// Licenses
// ---------------------------------------------------------------------------------------------

type SkuKey = "SPB" | "EXCHANGESTANDARD" | "POWER_BI_PRO" | "AAD_PREMIUM_P2" | "FLOW_FREE";

const skuPlans: Record<SkuKey, string[]> = {
  SPB: [
    "EXCHANGE_S_ENTERPRISE",
    "TEAMS1",
    "SHAREPOINTENTERPRISE",
    "OFFICESUBSCRIPTION",
    "AAD_PREMIUM",
    "INTUNE_A",
    "MFA_PREMIUM",
  ],
  EXCHANGESTANDARD: ["EXCHANGE_S_STANDARD", "MICROSOFTBOOKINGS"],
  POWER_BI_PRO: ["BI_AZURE_P2"],
  AAD_PREMIUM_P2: ["AAD_PREMIUM_P2", "MFA_PREMIUM"],
  FLOW_FREE: ["FLOW_P2_VIRAL"],
};

const skuId = (key: SkuKey) => guid(`sku:${key}`);

// ---------------------------------------------------------------------------------------------
// The roster
// ---------------------------------------------------------------------------------------------

type Kind = "active" | "inactive" | "never" | "disabled" | "guest" | "shared";
type Mfa = "passwordless" | "app" | "phone" | "none";

interface Seed {
  key: string;
  name: string;
  title: string;
  department: string;
  site: SiteId;
  manager?: string;
  skus: SkuKey[];
  kind: Kind;
  mfa: Mfa;
  admin?: boolean;
  hybrid?: boolean;
  // Days since the last successful sign-in, for accounts that went quiet.
  quietDays?: number;
  // Days since the account was created, when it is not simply old.
  createdDaysAgo?: number;
  // The license reached the account through a group rather than directly.
  byGroup?: boolean;
  // Guests only.
  guest?: { domain: string; state: "Accepted" | "PendingAcceptance" };
  legacyPhone?: boolean;
  apps?: AppKey[];
}

const lab = "Laboratory";
const roster: Seed[] = [
  {
    key: "admin",
    name: "Demo administrator",
    title: "IT Administrator",
    department: "IT",
    site: "madison",
    manager: "mbrandt",
    skus: ["SPB", "AAD_PREMIUM_P2"],
    kind: "active",
    mfa: "passwordless",
    admin: true,
    hybrid: true,
    apps: ["azure", "portal", "teams", "exchange"],
  },
  {
    key: "tvance",
    name: "Tobias Vance",
    title: "IT Support Technician",
    department: "IT",
    site: "madison",
    manager: "admin",
    skus: ["SPB", "AAD_PREMIUM_P2"],
    kind: "active",
    mfa: "none",
    admin: true,
    hybrid: true,
    apps: ["azure", "exchange", "teams", "portal"],
  },
  {
    key: "mbrandt",
    name: "Mateus Brandt",
    title: "General Manager",
    department: "Executive",
    site: "madison",
    skus: ["SPB", "AAD_PREMIUM_P2"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "isolheim",
    name: "Ingrid Solheim",
    title: "Laboratory Director",
    department: lab,
    site: "madison",
    manager: "mbrandt",
    skus: ["SPB"],
    kind: "active",
    mfa: "passwordless",
    hybrid: true,
  },
  {
    key: "dholloway",
    name: "Dagny Holloway",
    title: "Senior Analytical Chemist",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "korne",
    name: "Kasimir Orne",
    title: "Quality Assurance Manager",
    department: "Quality",
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "waris",
    name: "Wendell Aris",
    title: "Method Development Chemist",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "mquint",
    name: "Marisol Quint",
    title: "Laboratory Technician",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["EXCHANGESTANDARD"],
    kind: "active",
    mfa: "phone",
    hybrid: true,
    legacyPhone: true,
    apps: ["exchange", "teams"],
  },
  {
    key: "emarrow",
    name: "Ezekiel Marrow",
    title: "Laboratory Technician",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["EXCHANGESTANDARD"],
    kind: "active",
    mfa: "phone",
    hybrid: true,
    apps: ["exchange", "teams"],
  },
  {
    key: "pdessai",
    name: "Priyanka Dessai",
    title: "Chromatography Specialist",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "hbeck",
    name: "Halvard Beck",
    title: "Instrument Specialist",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "ofenwick",
    name: "Odalys Fenwick",
    title: "Sample Receiving Coordinator",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["EXCHANGESTANDARD"],
    kind: "active",
    mfa: "phone",
    hybrid: true,
    apps: ["exchange", "teams"],
  },
  {
    key: "ctillery",
    name: "Corwin Tillery",
    title: "Laboratory Technician",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["EXCHANGESTANDARD"],
    kind: "active",
    mfa: "phone",
    hybrid: true,
    legacyPhone: true,
    apps: ["exchange", "teams"],
  },
  {
    key: "jashdown",
    name: "Juniper Ashdown",
    title: "Research Scientist",
    department: "Research",
    site: "madison",
    manager: "isolheim",
    skus: ["SPB", "POWER_BI_PRO"],
    kind: "active",
    mfa: "app",
    apps: ["exchange", "teams", "sharepoint", "powerbi"],
  },
  {
    key: "lverhoeven",
    name: "Lucan Verhoeven",
    title: "Research Scientist",
    department: "Research",
    site: "madison",
    manager: "isolheim",
    skus: ["SPB", "POWER_BI_PRO"],
    kind: "active",
    mfa: "app",
    apps: ["exchange", "teams", "sharepoint", "powerbi"],
  },
  {
    key: "spetrakis",
    name: "Sunniva Petrakis",
    title: "Regulatory Affairs Specialist",
    department: "Quality",
    site: "madison",
    manager: "korne",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },
  {
    key: "akirov",
    name: "Annelise Kirov",
    title: "Finance Manager",
    department: "Finance",
    site: "madison",
    manager: "mbrandt",
    skus: ["SPB", "POWER_BI_PRO"],
    kind: "active",
    mfa: "app",
    hybrid: true,
    apps: ["exchange", "teams", "sharepoint", "powerbi"],
  },
  {
    key: "bokafor",
    name: "Bram Okafor",
    title: "Accounts Payable Clerk",
    department: "Finance",
    site: "madison",
    manager: "akirov",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "rteague",
    name: "Rosalind Teague",
    title: "HR Generalist",
    department: "People",
    site: "madison",
    manager: "mbrandt",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
    hybrid: true,
  },
  {
    key: "flowe",
    name: "Ferdinand Lowe",
    title: "Facilities Coordinator",
    department: "Operations",
    site: "madison",
    manager: "mbrandt",
    skus: ["EXCHANGESTANDARD"],
    kind: "active",
    mfa: "phone",
    hybrid: true,
    apps: ["exchange", "teams"],
  },
  {
    key: "yhalvorsen",
    name: "Yusra Halvorsen",
    title: "Client Services Manager",
    department: "Client Services",
    site: "minneapolis",
    manager: "mbrandt",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },
  {
    key: "dferrante",
    name: "Desmond Ferrante",
    title: "Account Executive",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },
  {
    key: "mcobb",
    name: "Mirela Cobb",
    title: "Client Services Representative",
    department: "Client Services",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },
  {
    key: "awhitlock",
    name: "Anders Whitlock",
    title: "Account Executive",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },
  {
    key: "npellegrino",
    name: "Noor Pellegrino",
    title: "Business Analyst",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB", "POWER_BI_PRO"],
    kind: "active",
    mfa: "app",
    apps: ["exchange", "teams", "sharepoint", "powerbi"],
  },
  {
    key: "timrie",
    name: "Tavish Imrie",
    title: "Sales Operations Analyst",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB", "POWER_BI_PRO"],
    kind: "active",
    mfa: "app",
    apps: ["exchange", "teams", "sharepoint", "powerbi"],
  },
  {
    key: "labara",
    name: "Linnea Abara",
    title: "Data Scientist",
    department: "Research",
    site: "denver",
    manager: "isolheim",
    skus: ["SPB", "POWER_BI_PRO"],
    kind: "active",
    mfa: "passwordless",
    apps: ["exchange", "teams", "sharepoint", "powerbi"],
  },
  {
    key: "csedgwick",
    name: "Cormac Sedgwick",
    title: "Software Developer",
    department: "Research",
    site: "raleigh",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "active",
    mfa: "passwordless",
  },
  {
    key: "zbrandvold",
    name: "Zofia Brandvold",
    title: "Technical Writer",
    department: "Quality",
    site: "portland",
    manager: "korne",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },
  {
    key: "icalloway",
    name: "Idris Calloway",
    title: "Field Service Engineer",
    department: "Operations",
    site: "austin",
    manager: "mbrandt",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },
  {
    key: "blindqvist",
    name: "Beatrix Lindqvist",
    title: "Proposal Writer",
    department: "Sales",
    site: "raleigh",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "active",
    mfa: "app",
  },

  // Accounts that went quiet, oldest last sign-in last.
  {
    key: "pstrand",
    name: "Philippa Strand",
    title: "Account Executive",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 8,
    mfa: "app",
  },
  {
    key: "ovenn",
    name: "Oskar Venn",
    title: "Laboratory Technician",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["EXCHANGESTANDARD"],
    kind: "inactive",
    quietDays: 14,
    mfa: "phone",
    hybrid: true,
  },
  {
    key: "rdoyle",
    name: "Ravenna Doyle",
    title: "Marketing Coordinator",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 35,
    mfa: "app",
  },
  {
    key: "ewragg",
    name: "Elias Wragg",
    title: "Summer Intern",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 52,
    mfa: "app",
    createdDaysAgo: 160,
  },
  {
    key: "cdunmore",
    name: "Calla Dunmore",
    title: "Sales Operations Analyst",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 70,
    mfa: "none",
  },
  {
    key: "jfennimore",
    name: "Jory Fennimore",
    title: "Contract Chemist",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 91,
    mfa: "app",
    createdDaysAgo: 300,
  },
  {
    key: "msterling",
    name: "Magnus Sterling",
    title: "Laboratory Technician",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["EXCHANGESTANDARD"],
    kind: "inactive",
    quietDays: 100,
    mfa: "phone",
    hybrid: true,
  },
  {
    key: "trooke",
    name: "Tamsin Rooke",
    title: "Quality Specialist",
    department: "Quality",
    site: "madison",
    manager: "korne",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 130,
    mfa: "app",
    byGroup: true,
    hybrid: true,
  },
  {
    key: "agray",
    name: "Anselm Gray",
    title: "Contractor",
    department: "Operations",
    site: "madison",
    manager: "mbrandt",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 160,
    mfa: "none",
    createdDaysAgo: 420,
  },
  {
    key: "dokoro",
    name: "Delphine Okoro",
    title: "Customer Success Lead",
    department: "Client Services",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 200,
    mfa: "app",
  },
  {
    key: "rpemberton",
    name: "Rufus Pemberton",
    title: "Seasonal Technician",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["EXCHANGESTANDARD"],
    kind: "inactive",
    quietDays: 290,
    mfa: "phone",
    createdDaysAgo: 520,
  },
  {
    key: "iquill",
    name: "Isadora Quill",
    title: "Former Consultant",
    department: "Operations",
    site: "madison",
    manager: "mbrandt",
    skus: ["SPB"],
    kind: "inactive",
    quietDays: 400,
    mfa: "none",
    createdDaysAgo: 700,
  },

  // Never signed in.
  {
    key: "balder",
    name: "Brigid Alder",
    title: "Laboratory Technician (starts soon)",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "never",
    createdDaysAgo: 10,
    mfa: "none",
  },
  {
    key: "qhale",
    name: "Quentin Hale",
    title: "Contractor (onboarding)",
    department: "Operations",
    site: "madison",
    manager: "mbrandt",
    skus: ["SPB"],
    kind: "never",
    createdDaysAgo: 60,
    mfa: "none",
  },
  {
    key: "kiosk",
    name: "Lab Training Kiosk",
    title: "Training account",
    department: lab,
    site: "madison",
    skus: [],
    kind: "never",
    createdDaysAgo: 150,
    mfa: "none",
  },

  // Disabled, with licenses still assigned.
  {
    key: "hmoss",
    name: "Harriet Moss",
    title: "Former Chemist",
    department: lab,
    site: "madison",
    manager: "isolheim",
    skus: ["SPB"],
    kind: "disabled",
    quietDays: 180,
    mfa: "app",
    hybrid: true,
  },
  {
    key: "lvass",
    name: "Leopold Vass",
    title: "Former Account Executive",
    department: "Sales",
    site: "minneapolis",
    manager: "yhalvorsen",
    skus: ["SPB"],
    kind: "disabled",
    quietDays: 250,
    mfa: "app",
  },

  // A shared mailbox: it never signs in and is blocked from doing so.
  {
    key: "labresults",
    name: "Lab Results (shared)",
    title: "Shared mailbox",
    department: lab,
    site: "madison",
    skus: [],
    kind: "shared",
    mfa: "none",
    hybrid: true,
  },

  // Guests.
  {
    key: "awhitcombe",
    name: "Alder Whitcombe",
    title: "Quality auditor",
    department: "External",
    site: "madison",
    skus: [],
    kind: "guest",
    quietDays: 6,
    mfa: "app",
    guest: { domain: "harborline.example", state: "Accepted" },
    createdDaysAgo: 120,
  },
  {
    key: "mtadesse",
    name: "Mireille Tadesse",
    title: "Consulting chemist",
    department: "External",
    site: "madison",
    skus: [],
    kind: "guest",
    quietDays: 150,
    mfa: "app",
    guest: { domain: "quillmark.example", state: "Accepted" },
    createdDaysAgo: 300,
  },
  {
    key: "crousseau",
    name: "Callum Rousseau",
    title: "Equipment vendor",
    department: "External",
    site: "madison",
    skus: [],
    kind: "guest",
    mfa: "none",
    guest: { domain: "northgate-labs.example", state: "PendingAcceptance" },
    createdDaysAgo: 45,
  },
  {
    key: "sokonkwo",
    name: "Sabine Okonkwo-Reyes",
    title: "Accounts contact",
    department: "External",
    site: "madison",
    skus: [],
    kind: "guest",
    mfa: "none",
    guest: { domain: "quillmark.example", state: "PendingAcceptance" },
    createdDaysAgo: 9,
  },
];

const capitalise = (value: string) => `${value.charAt(0).toUpperCase()}${value.slice(1)}`;
const slug = (name: string) =>
  name
    .toLowerCase()
    .replace(/\s*\(.*\)$/, "")
    .replace(/[^a-z0-9]+/g, ".")
    .replace(/^\.+|\.+$/g, "");

// ---------------------------------------------------------------------------------------------
// The directory
// ---------------------------------------------------------------------------------------------

export interface DirectoryUser {
  id: string;
  displayName: string;
  givenName: string;
  surname: string;
  userPrincipalName: string;
  mail: string;
  jobTitle: string;
  department: string;
  companyName: string;
  officeLocation: string;
  employeeId: string | null;
  userType: "Member" | "Guest";
  accountEnabled: boolean;
  createdDateTime: string;
  lastPasswordChangeDateTime: string | null;
  onPremisesSyncEnabled: boolean | null;
  usageLocation: string;
  city: string;
  state: string;
  country: string;
  mobilePhone: string | null;
  businessPhones: string[];
  externalUserState: string | null;
  externalUserStateChangeDateTime: string | null;
  assignedLicenses: { skuId: string; disabledPlans: string[] }[];
  licenseAssignmentStates: {
    skuId: string;
    assignedByGroup: string | null;
    state: string;
    error: string;
  }[];
  signInActivity: {
    lastSignInDateTime: string | null;
    lastNonInteractiveSignInDateTime: string | null;
    lastSuccessfulSignInDateTime: string | null;
  };
}

interface Member {
  seed: Seed;
  user: DirectoryUser;
  managerId: string | null;
  devices: { laptop: Device; phone: Device; legacy: Device | null };
}

export interface SignIn {
  id: string;
  createdDateTime: string;
  userId: string;
  userPrincipalName: string;
  userDisplayName: string;
  appId: string;
  appDisplayName: string;
  resourceDisplayName: string;
  ipAddress: string;
  clientAppUsed: string;
  conditionalAccessStatus: string;
  isInteractive: boolean;
  riskLevelDuringSignIn: string;
  riskLevelAggregated: string;
  riskState: string;
  riskEventTypes_v2: string[];
  status: { errorCode: number; failureReason: string; additionalDetails: string | null };
  deviceDetail: {
    deviceId: string;
    operatingSystem: string;
    browser: string;
    isCompliant: boolean | null;
    isManaged: boolean | null;
    trustType: string | null;
  };
  location: {
    city: string;
    state: string;
    countryOrRegion: string;
    geoCoordinates: { latitude: number; longitude: number };
  };
}

export interface SyntheticTenant {
  now: Date;
  users: DirectoryUser[];
  managers: Map<string, string | null>;
  signIns: SignIn[];
  skus: ReturnType<typeof buildSkus>;
  registrations: ReturnType<typeof buildRegistrations>;
  // Every address a synthetic response may contain, for the recorder's closure check.
  addresses: Set<string>;
  // Accounts the planted stories are about, by key.
  stories: Record<string, string>;
}

function deviceFor(rand: Rand, key: string, kind: "laptop" | "phone" | "legacy"): Device {
  if (kind === "laptop") {
    const mac = chance(rand, 0.25);
    return {
      id: guid(`device:${key}:laptop`),
      os: mac ? "macOS 15.6" : "Windows 11",
      browser: mac
        ? pick(rand, ["Safari 18.6", "Chrome 141.0.7390"])
        : pick(rand, ["Edge 141.0.3537", "Chrome 141.0.7390"]),
      managed: true,
      compliant: true,
      trustType: mac ? "Workplace" : "AzureAd",
    };
  }
  if (kind === "phone") {
    const ios = chance(rand, 0.6);
    return {
      id: guid(`device:${key}:phone`),
      os: ios ? "iOS 18.6" : "Android 15",
      browser: ios ? "Mobile Safari 18.6" : "Chrome Mobile 141.0.7390",
      managed: chance(rand, 0.5),
      compliant: chance(rand, 0.5) ? true : null,
      trustType: null,
    };
  }
  return {
    id: guid(`device:${key}:legacy`),
    os: "Android 12",
    browser: "",
    managed: false,
    compliant: null,
    trustType: null,
  };
}

function buildMembers(now: Date): Member[] {
  const ids = new Map(roster.map((seed) => [seed.key, guid(`user:${seed.key}`)]));
  const t = now.getTime();
  return roster.map((seed, index) => {
    const rand = seeded(`member:${seed.key}`);
    const site = sites[seed.site];
    const [given = "", ...rest] = seed.name.split(" ");
    const surname = rest.join(" ");
    const local = seed.key === "admin" ? "admin" : slug(seed.name);
    const guest = seed.guest;
    const home = seed.department === lab ? RESEARCH_DOMAIN : DOMAIN;
    const upn = guest ? `${local}_${guest.domain}#EXT#@${DOMAIN}` : `${local}@${home}`;
    const mail = guest ? `${local}@${guest.domain}` : upn;
    const created =
      seed.createdDaysAgo !== undefined
        ? t - seed.createdDaysAgo * DAY - between(rand, 0, 20) * HOUR
        : t - between(rand, 300, 1_400) * DAY - between(rand, 0, 20) * HOUR;
    const lastPassword = t - between(rand, 5, 80) * DAY - between(rand, 0, 20) * HOUR;
    const licenses = seed.skus.slice();
    // Power Automate's free plan is handed out with the main license to part of the company.
    if (seed.skus.includes("SPB") && index % 3 !== 0 && seed.kind !== "never")
      licenses.push("FLOW_FREE");
    const groupId = guid("group:licensing");
    const user: DirectoryUser = {
      id: ids.get(seed.key) as string,
      displayName: seed.name,
      givenName: given,
      surname,
      userPrincipalName: upn,
      mail,
      jobTitle: seed.title,
      department: seed.department,
      companyName: guest ? capitalise(guest.domain.split(".")[0] ?? "External") : ORGANISATION,
      officeLocation: guest ? "External" : site.office,
      employeeId: guest || seed.kind === "shared" ? null : `LA-${String(1_000 + index)}`,
      userType: guest ? "Guest" : "Member",
      accountEnabled: seed.kind !== "disabled" && seed.kind !== "shared",
      createdDateTime: iso(created),
      lastPasswordChangeDateTime: guest || seed.kind === "shared" ? null : iso(lastPassword),
      onPremisesSyncEnabled: seed.hybrid ? true : null,
      usageLocation: "US",
      city: site.city,
      state: site.state,
      country: "United States",
      mobilePhone:
        guest || seed.kind === "shared"
          ? null
          : `+1 ${seed.site === "madison" ? "608" : "612"} 555 01${String(10 + (index % 80))}`,
      businessPhones:
        seed.kind === "shared" || guest ? [] : [`+1 608 555 01${String(index % 10)}0`],
      externalUserState: guest ? guest.state : null,
      externalUserStateChangeDateTime: guest
        ? iso(created + (guest.state === "Accepted" ? 2 * DAY : 0))
        : null,
      assignedLicenses: licenses.map((key) => ({ skuId: skuId(key as SkuKey), disabledPlans: [] })),
      licenseAssignmentStates: licenses.map((key) => ({
        skuId: skuId(key as SkuKey),
        assignedByGroup: seed.byGroup && key === "SPB" ? groupId : null,
        state: "Active",
        error: "None",
      })),
      signInActivity: {
        lastSignInDateTime: null,
        lastNonInteractiveSignInDateTime: null,
        lastSuccessfulSignInDateTime: null,
      },
    };
    return {
      seed,
      user,
      managerId: seed.manager ? (ids.get(seed.manager) ?? null) : null,
      devices: {
        laptop: deviceFor(rand, seed.key, "laptop"),
        phone: deviceFor(rand, seed.key, "phone"),
        legacy: seed.legacyPhone ? deviceFor(rand, seed.key, "legacy") : null,
      },
    };
  });
}

// ---------------------------------------------------------------------------------------------
// Sign-ins
// ---------------------------------------------------------------------------------------------

const appWeights: [AppKey, number][] = [
  ["exchange", 30],
  ["teams", 30],
  ["sharepoint", 15],
  ["portal", 10],
  ["onedrive", 10],
  ["powerbi", 5],
];

function chooseApp(rand: Rand, member: Member): AppKey {
  const allowed = member.seed.apps;
  const pool = allowed
    ? allowed.map(
        (key) =>
          [key, appWeights.find(([candidate]) => candidate === key)?.[1] ?? 20] as [AppKey, number],
      )
    : appWeights.filter(([key]) => key !== "powerbi" && key !== "azure");
  const total = pool.reduce((sum, [, weight]) => sum + weight, 0);
  let roll = rand() * total;
  for (const [key, weight] of pool) {
    roll -= weight;
    if (roll <= 0) return key;
  }
  return pool[0]?.[0] ?? "exchange";
}

interface EventSpec {
  member: Member;
  at: number;
  ip: string;
  place: Place;
  device: {
    id: string;
    os: string;
    browser: string;
    managed: boolean | null;
    compliant: boolean | null;
    trustType: string | null;
  };
  app: AppKey;
  client?: string;
  errorCode: number;
  reason?: string;
  details?: string | null;
  ca?: string;
  label: string;
}

function makeSignIn(spec: EventSpec): SignIn {
  const app = apps[spec.app];
  const success = spec.errorCode === 0;
  return {
    id: guid(`signin:${spec.member.seed.key}:${spec.label}:${spec.at}`),
    createdDateTime: iso(spec.at),
    userId: spec.member.user.id,
    userPrincipalName: spec.member.user.userPrincipalName,
    userDisplayName: spec.member.user.displayName,
    appId: guid(`app:${app.name}`),
    appDisplayName: app.name,
    resourceDisplayName: app.resource,
    ipAddress: spec.ip,
    clientAppUsed:
      spec.client ?? (spec.device.browser ? "Browser" : "Mobile Apps and Desktop clients"),
    conditionalAccessStatus: spec.ca ?? (success ? "success" : "notApplied"),
    isInteractive: true,
    riskLevelDuringSignIn: "none",
    riskLevelAggregated: "none",
    riskState: "none",
    riskEventTypes_v2: [],
    status: {
      errorCode: spec.errorCode,
      failureReason: success ? "Other." : (spec.reason ?? "Other."),
      additionalDetails: spec.details ?? null,
    },
    deviceDetail: {
      deviceId: spec.device.id,
      operatingSystem: spec.device.os,
      browser: spec.device.browser,
      isCompliant: spec.device.compliant,
      isManaged: spec.device.managed,
      trustType: spec.device.trustType,
    },
    location: {
      city: spec.place.city,
      state: spec.place.state,
      countryOrRegion: spec.place.country,
      geoCoordinates: { latitude: spec.place.latitude, longitude: spec.place.longitude },
    },
  };
}

const interrupt = {
  errorCode: 50074,
  reason: "Strong Authentication is required.",
  details: "MFA required in Microsoft Entra ID",
};
const typo = {
  errorCode: 50126,
  reason: "Invalid username or password or Invalid on-premise username or password.",
};

// A weekday of ordinary office sign-ins for one person: a morning one, usually a midday one and
// sometimes an afternoon one, from the same few places and devices every day.
function ordinarySignIns(member: Member, now: number, rand: Rand): SignIn[] {
  const seed = member.seed;
  const site = sites[seed.site];
  const events: SignIn[] = [];
  const today = startOfDay(now);
  const slots: [number, number, number][] = [
    [7.8, 9.6, 1],
    [12, 13.5, 0.65],
    [14.75, 16.75, seed.admin ? 0.75 : 0.3],
  ];
  for (let offset = 30; offset >= 0; offset -= 1) {
    const day = today - offset * DAY;
    const weekday = new Date(day).getUTCDay();
    const working = weekday >= 1 && weekday <= 5 && !chance(rand, 0.07);
    if (!working && !chance(rand, 0.03)) continue;
    let first = true;
    for (const [from, to, probability] of slots) {
      if (!chance(rand, probability)) continue;
      const hour = from + rand() * (to - from);
      const at = day + (hour - site.utcOffset) * HOUR + between(rand, 0, 59) * 1_000;
      const useLaptop = chance(rand, 0.85);
      const device = useLaptop ? member.devices.laptop : member.devices.phone;
      const ip = pick(rand, site.ips);
      const app = seed.admin && chance(rand, 0.3) ? "azure" : chooseApp(rand, member);
      const base = { member, ip, place: site, device, app } satisfies Omit<
        EventSpec,
        "at" | "errorCode" | "label"
      >;
      if (first && chance(rand, 0.35))
        events.push(makeSignIn({ ...base, at: at - 15_000, ...interrupt, label: "mfa" }));
      if (chance(rand, 0.012))
        events.push(makeSignIn({ ...base, at: at - 45_000, ...typo, label: "typo" }));
      events.push(makeSignIn({ ...base, at, errorCode: 0, label: "ok" }));
      first = false;
    }
    // Old Android phones that still sync mail over ActiveSync.
    if (member.devices.legacy && working && chance(rand, 0.5)) {
      const at =
        day + (between(rand, 7, 17) - site.utcOffset) * HOUR + between(rand, 0, 3_599) * 1_000;
      events.push(
        makeSignIn({
          member,
          at,
          ip: pick(rand, site.ips),
          place: site,
          device: member.devices.legacy,
          app: "exchange",
          client: "Exchange ActiveSync",
          errorCode: 0,
          label: "legacy",
        }),
      );
    }
  }
  return events;
}

interface Stories {
  ids: Record<string, string>;
  events: SignIn[];
  // Ordinary sign-ins to drop for the accounts the stories are about.
  suppress: ((member: Member, event: SignIn) => boolean)[];
}

function plantStories(members: Map<string, Member>, now: number): Stories {
  const member = (key: string) => {
    const found = members.get(key);
    if (!found) throw new Error(`Unknown synthetic account ${key}.`);
    return found;
  };
  const events: SignIn[] = [];
  const suppress: Stories["suppress"] = [];
  const day = (daysAgo: number, hour: number, minute = 0, second = 0) =>
    startOfDay(now) - daysAgo * DAY + hour * HOUR + minute * MINUTE + second * 1_000;
  const unmanaged = (key: string, os: string, browser: string) => ({
    id: guid(`device:${key}`),
    os,
    browser,
    managed: false as boolean | null,
    compliant: false as boolean | null,
    trustType: null,
  });

  // 1. Impossible travel: a success from Bucharest on an unknown laptop about an hour after the
  // same account signed in at its Madison laboratory.
  const traveler = member("dholloway");
  const madison = sites.madison;
  events.push(
    makeSignIn({
      member: traveler,
      at: now - 105 * MINUTE,
      ip: madison.ips[0] as string,
      place: madison,
      device: traveler.devices.laptop,
      app: "teams",
      errorCode: 0,
      label: "home-base",
    }),
    makeSignIn({
      member: traveler,
      at: now - 41 * MINUTE - 20_000,
      ip: TRAVEL_IP,
      place: bucharest,
      device: unmanaged("travel", "Windows 10", "Firefox 142.0"),
      app: "exchange",
      ...interrupt,
      label: "travel-mfa",
    }),
    makeSignIn({
      member: traveler,
      at: now - 41 * MINUTE,
      ip: TRAVEL_IP,
      place: bucharest,
      device: unmanaged("travel", "Windows 10", "Firefox 142.0"),
      app: "exchange",
      errorCode: 0,
      label: "travel",
    }),
  );
  suppress.push(
    (who, event) =>
      who.seed.key === "dholloway" && Date.parse(event.createdDateTime) > now - 110 * MINUTE,
  );

  // 2. Password spray: one address tries one password against seven accounts, eight minutes apart.
  const victims = ["bokafor", "rteague", "hbeck", "spetrakis", "akirov", "dferrante", "mcobb"];
  for (const [index, key] of victims.entries())
    events.push(
      makeSignIn({
        member: member(key),
        at: day(11, 8, 5 + index * 8, 12),
        ip: SPRAY_IP,
        place: frankfurt,
        device: {
          id: "",
          os: "Linux",
          browser: "",
          managed: null,
          compliant: null,
          trustType: null,
        },
        app: "exchange",
        client: "Other clients",
        errorCode: 50126,
        reason: typo.reason,
        label: "spray",
      }),
    );

  // 3. Failed MFA prompts: someone with the password keeps trying until the account holder
  // would have to approve.
  const target = member("korne");
  for (let index = 0; index < 5; index += 1)
    events.push(
      makeSignIn({
        member: target,
        at: day(4, 8, 12, index * 90),
        ip: MFA_ATTACKER_IP,
        place: ashburn,
        device: unmanaged("mfa-attacker", "Windows 10", "Chrome 140.0.7339"),
        app: "portal",
        errorCode: 500121,
        reason: "Authentication failed during strong authentication request.",
        label: `mfa-${index}`,
      }),
    );

  // 4. A success right after a run of failures, from the account's own laptop at home.
  const forgetful = member("waris");
  const home = { ...sites.madison, ips: [WENDELL_HOME_IP] };
  for (let index = 0; index < 5; index += 1)
    events.push(
      makeSignIn({
        member: forgetful,
        at: day(2, 13, 41, index * 70),
        ip: WENDELL_HOME_IP,
        place: home,
        device: forgetful.devices.laptop,
        app: "portal",
        ...typo,
        label: `fail-${index}`,
      }),
    );
  events.push(
    makeSignIn({
      member: forgetful,
      at: day(2, 13, 49, 10),
      ip: WENDELL_HOME_IP,
      place: home,
      device: forgetful.devices.laptop,
      app: "portal",
      errorCode: 0,
      label: "after-failures",
    }),
  );
  suppress.push((who, event) => {
    if (who.seed.key !== "waris") return false;
    const at = Date.parse(event.createdDateTime);
    return at > day(2, 13, 0) && at < day(2, 15, 0);
  });

  // 5. Conditional Access turning away unmanaged personal laptops.
  for (const [key, daysAgo, minute] of [
    ["csedgwick", 17, 20],
    ["csedgwick", 17, 31],
    ["zbrandvold", 9, 4],
    ["icalloway", 3, 52],
  ] as const) {
    const who = member(key);
    const site = sites[who.seed.site];
    events.push(
      makeSignIn({
        member: who,
        at: day(daysAgo, 16, minute),
        ip: site.ips[0] as string,
        place: site,
        device: unmanaged(`personal-${key}`, "Linux", "Firefox 142.0"),
        app: "sharepoint",
        errorCode: 53003,
        reason:
          "Access has been blocked by Conditional Access policies. The access policy does not allow token issuance.",
        ca: "failure",
        label: `blocked-${daysAgo}`,
      }),
    );
  }

  // 6. A few people who are signed in right now, whatever the hour of the recording.
  for (const [key, minutesAgo] of [
    ["admin", 11],
    ["isolheim", 27],
    ["pdessai", 34],
    ["yhalvorsen", 19],
    ["jashdown", 52],
  ] as const) {
    const who = member(key);
    const site = sites[who.seed.site];
    events.push(
      makeSignIn({
        member: who,
        at: now - minutesAgo * MINUTE,
        ip: site.ips[0] as string,
        place: site,
        device: who.devices.laptop,
        app: key === "admin" ? "azure" : "teams",
        errorCode: 0,
        label: "now",
      }),
    );
  }

  return {
    ids: {
      traveler: traveler.user.id,
      mfaTarget: target.user.id,
      forgetful: forgetful.user.id,
    },
    events,
    suppress,
  };
}

// ---------------------------------------------------------------------------------------------
// Skus and authentication methods
// ---------------------------------------------------------------------------------------------

function buildSkus(members: Member[]) {
  const consumed = (key: SkuKey) =>
    members.filter((member) =>
      member.user.assignedLicenses.some((license) => license.skuId === skuId(key)),
    ).length;
  const sku = (
    key: SkuKey,
    units: { enabled: number; warning?: number; suspended?: number; lockedOut?: number },
    status: string,
  ) => ({
    id: `${SYNTHETIC_TENANT_ID}_${skuId(key)}`,
    skuId: skuId(key),
    skuPartNumber: key,
    appliesTo: "User",
    capabilityStatus: status,
    consumedUnits: consumed(key),
    prepaidUnits: {
      enabled: units.enabled,
      suspended: units.suspended ?? 0,
      warning: units.warning ?? 0,
      lockedOut: units.lockedOut ?? 0,
    },
    servicePlans: skuPlans[key].map((name) => ({
      servicePlanId: guid(`plan:${name}`),
      servicePlanName: name,
      provisioningStatus: "Success",
      appliesTo: "User",
    })),
  });
  return [
    // No seats left: every seat bought is assigned.
    sku("SPB", { enabled: consumed("SPB") }, "Enabled"),
    // Three seats to spare.
    sku("EXCHANGESTANDARD", { enabled: consumed("EXCHANGESTANDARD") + 3 }, "Enabled"),
    // Expired subscription still inside its grace period.
    sku("POWER_BI_PRO", { enabled: 0, warning: consumed("POWER_BI_PRO") + 2 }, "Warning"),
    sku("AAD_PREMIUM_P2", { enabled: consumed("AAD_PREMIUM_P2") + 2 }, "Enabled"),
    // A free product Microsoft provisions with a very large seat count.
    sku("FLOW_FREE", { enabled: 10_000 }, "Enabled"),
  ];
}

const methodSets: Record<Mfa, string[][]> = {
  passwordless: [
    [
      "fido2SecurityKey",
      "microsoftAuthenticatorPasswordless",
      "microsoftAuthenticatorPush",
      "softwareOneTimePasscode",
    ],
    ["windowsHelloForBusiness", "microsoftAuthenticatorPush", "mobilePhone"],
  ],
  app: [
    ["microsoftAuthenticatorPush", "softwareOneTimePasscode", "mobilePhone"],
    ["microsoftAuthenticatorPush", "softwareOneTimePasscode"],
    ["microsoftAuthenticatorPush", "email"],
  ],
  phone: [["mobilePhone"], ["mobilePhone", "email"]],
  none: [[], ["email"]],
};

function buildRegistrations(members: Member[], now: number) {
  return members
    .filter((member) => member.user.accountEnabled)
    .map((member, index) => {
      const rand = seeded(`mfa:${member.seed.key}`);
      const methods = pick(rand, methodSets[member.seed.mfa]);
      const registered = member.seed.mfa !== "none";
      const passwordless = member.seed.mfa === "passwordless";
      const preferred = passwordless
        ? "push"
        : registered
          ? methods.includes("mobilePhone") && member.seed.mfa === "phone"
            ? "sms"
            : "push"
          : "none";
      return {
        id: member.user.id,
        userPrincipalName: member.user.userPrincipalName,
        userDisplayName: member.user.displayName,
        userType: member.user.userType === "Guest" ? "guest" : "member",
        isAdmin: Boolean(member.seed.admin),
        isMfaRegistered: registered,
        isMfaCapable: registered,
        isPasswordlessCapable: passwordless,
        isSsprRegistered: member.seed.mfa !== "none" || index % 4 === 0,
        isSsprEnabled: member.user.userType === "Member",
        isSsprCapable: registered,
        methodsRegistered: methods,
        isSystemPreferredAuthenticationMethodEnabled: registered && member.seed.mfa !== "phone",
        systemPreferredAuthenticationMethods:
          registered && member.seed.mfa !== "phone" ? ["push"] : [],
        userPreferredMethodForSecondaryAuthentication: preferred,
        lastUpdatedDateTime: iso(now - between(rand, 1, 12) * HOUR - between(rand, 0, 50) * MINUTE),
      };
    });
}

// ---------------------------------------------------------------------------------------------
// Assembling the tenant
// ---------------------------------------------------------------------------------------------

export function buildSyntheticTenant(now: Date): SyntheticTenant {
  const t = now.getTime();
  const members = buildMembers(now);
  const byKey = new Map(members.map((member) => [member.seed.key, member]));
  const stories = plantStories(byKey, t);

  // Ordinary sign-ins, for the accounts that sign in within the 30 days Microsoft keeps.
  const signIns: SignIn[] = [];
  for (const member of members) {
    const seed = member.seed;
    const signsIn =
      seed.kind === "active" ||
      (seed.kind === "inactive" && (seed.quietDays ?? 999) < 30) ||
      (seed.kind === "guest" && (seed.quietDays ?? 999) < 30);
    if (!signsIn) continue;
    const rand = seeded(`signins:${seed.key}`);
    // Accounts that went quiet recently stop signing in at their last-seen time.
    const lastSeen =
      seed.quietDays !== undefined ? t - seed.quietDays * DAY - between(rand, 1, 20) * HOUR : t;
    const ordinary =
      seed.kind === "guest"
        ? [
            makeSignIn({
              member,
              at: lastSeen - 30 * HOUR,
              ip: sites.madison.ips[0] as string,
              place: sites.madison,
              device: member.devices.laptop,
              app: "sharepoint",
              errorCode: 0,
              label: "guest-1",
            }),
            makeSignIn({
              member,
              at: lastSeen - 4 * DAY - HOUR,
              ip: sites.madison.ips[0] as string,
              place: sites.madison,
              device: member.devices.laptop,
              app: "sharepoint",
              errorCode: 0,
              label: "guest-2",
            }),
            makeSignIn({
              member,
              at: lastSeen,
              ip: sites.madison.ips[0] as string,
              place: sites.madison,
              device: member.devices.laptop,
              app: "sharepoint",
              errorCode: 0,
              label: "guest-3",
            }),
          ]
        : ordinarySignIns(member, t, rand);
    // Accounts that went quiet show their last sign-in at exactly that time.
    if (seed.kind === "inactive")
      ordinary.push(
        makeSignIn({
          member,
          at: lastSeen,
          ip: sites[seed.site].ips[0] as string,
          place: sites[seed.site],
          device: member.devices.laptop,
          app: "teams",
          errorCode: 0,
          label: "last",
        }),
      );
    for (const event of ordinary) {
      const at = Date.parse(event.createdDateTime);
      if (at > lastSeen || at > t - 2 * MINUTE || at < t - 30 * DAY + 5 * MINUTE) continue;
      if (stories.suppress.some((rule) => rule(member, event))) continue;
      signIns.push(event);
    }
  }
  for (const event of stories.events) {
    const at = Date.parse(event.createdDateTime);
    if (at <= t - 2 * MINUTE && at >= t - 30 * DAY + 5 * MINUTE) signIns.push(event);
  }
  signIns.sort(
    (a, b) => a.createdDateTime.localeCompare(b.createdDateTime) || a.id.localeCompare(b.id),
  );

  // Directory last-sign-in dates follow what the log shows, or the account's age of silence.
  const latest = new Map<string, { interactive: number; success: number }>();
  for (const event of signIns) {
    const at = Date.parse(event.createdDateTime);
    const entry = latest.get(event.userId) ?? { interactive: 0, success: 0 };
    entry.interactive = Math.max(entry.interactive, at);
    if (event.status.errorCode === 0) entry.success = Math.max(entry.success, at);
    latest.set(event.userId, entry);
  }
  for (const member of members) {
    const rand = seeded(`activity:${member.seed.key}`);
    const seen = latest.get(member.user.id);
    const quiet = member.seed.quietDays;
    if (seen && seen.success > 0) {
      // Background token refreshes follow the last interactive sign-in: by a few hours for
      // accounts in daily use, by minutes for accounts that went quiet.
      const lag = quiet === undefined ? between(rand, 0, 5 * 60) : between(rand, 0, 30);
      const background = Math.min(t - 5 * MINUTE, seen.success + lag * MINUTE);
      member.user.signInActivity = {
        lastSignInDateTime: iso(seen.interactive),
        lastNonInteractiveSignInDateTime: iso(background),
        lastSuccessfulSignInDateTime: iso(Math.max(seen.success, background)),
      };
    } else if (quiet !== undefined && member.seed.kind !== "never") {
      const at = t - quiet * DAY - between(rand, 1, 20) * HOUR;
      const background = at + between(rand, 0, 30) * MINUTE;
      member.user.signInActivity = {
        lastSignInDateTime: iso(at),
        lastNonInteractiveSignInDateTime: iso(background),
        lastSuccessfulSignInDateTime: iso(background),
      };
    }
  }

  const users = members.map((member) => member.user);
  const addresses = new Set<string>();
  for (const user of users) {
    addresses.add(user.userPrincipalName.toLowerCase());
    addresses.add(user.mail.toLowerCase());
  }
  return {
    now,
    users,
    managers: new Map(members.map((member) => [member.user.id, member.managerId])),
    signIns,
    skus: buildSkus(members),
    registrations: buildRegistrations(members, t),
    addresses,
    stories: stories.ids,
  };
}

// ---------------------------------------------------------------------------------------------
// Service health and Message center
// ---------------------------------------------------------------------------------------------

function buildCommunications(t: number) {
  const overviews = [
    ["Exchange", "Exchange Online", "serviceDegradation"],
    ["Intune", "Microsoft Intune", "serviceOperational"],
    ["O365Client", "Microsoft 365 suite", "serviceOperational"],
    ["AzureActiveDirectory", "Microsoft Entra", "serviceOperational"],
    ["MicrosoftTeams", "Microsoft Teams", "serviceOperational"],
    ["OneDriveForBusiness", "OneDrive for Business", "serviceOperational"],
    ["Planner", "Planner", "serviceOperational"],
    ["PowerBIcom", "Power BI", "serviceOperational"],
    ["SharePoint", "SharePoint Online", "serviceOperational"],
  ].map(([id, service, status]) => ({ id, service, status }));

  const html = (lines: string[]) => lines.map((line) => `<p>${line}</p>`).join("");
  const issues = [
    {
      id: "EX900417",
      title: "Some users may be delayed when sending or receiving email in Exchange Online",
      service: "Exchange Online",
      feature: "E-Mail and calendar access",
      featureGroup: "E-Mail and calendar access",
      impactDescription:
        "Affected users may see delays of up to 20 minutes when sending or receiving email.",
      classification: "incident",
      status: "serviceDegradation",
      origin: "microsoft",
      isResolved: false,
      startDateTime: iso(t - 5 * HOUR),
      endDateTime: null,
      lastModifiedDateTime: iso(t - 35 * MINUTE),
      posts: [
        {
          createdDateTime: iso(t - 4 * HOUR - 50 * MINUTE),
          postType: "regular",
          description: {
            contentType: "html",
            content: html([
              "Title: Some users may be delayed when sending or receiving email in Exchange Online",
              "User impact: Affected users may see delays of up to 20 minutes when sending or receiving email.",
              "Current status: We are looking into a report of delayed mail delivery.",
            ]),
          },
        },
        {
          createdDateTime: iso(t - 35 * MINUTE),
          postType: "regular",
          description: {
            contentType: "html",
            content: html([
              "Title: Some users may be delayed when sending or receiving email in Exchange Online",
              "User impact: Affected users may see delays of up to 20 minutes when sending or receiving email.",
              "Current status: We found that one part of the mail flow infrastructure is processing messages more slowly than expected. We are moving traffic away from it and watching delivery times recover.",
              "Scope of impact: Any user served by the affected infrastructure may be impacted.",
              "Next update by: within 60 minutes",
            ]),
          },
        },
      ],
    },
    {
      id: "SP900233",
      title: "Some users may see slow page loads in SharePoint Online",
      service: "SharePoint Online",
      feature: "SharePoint Features",
      featureGroup: "SharePoint Features",
      impactDescription: "Some users saw SharePoint pages take longer than usual to open.",
      classification: "advisory",
      status: "serviceRestored",
      origin: "microsoft",
      isResolved: true,
      startDateTime: iso(t - 2 * DAY - 6 * HOUR),
      endDateTime: iso(t - 2 * DAY - 2 * HOUR),
      lastModifiedDateTime: iso(t - 2 * DAY - 90 * MINUTE),
      posts: [
        {
          createdDateTime: iso(t - 2 * DAY - 90 * MINUTE),
          postType: "regular",
          description: {
            contentType: "html",
            content: html([
              "Title: Some users may see slow page loads in SharePoint Online",
              "User impact: Some users saw SharePoint pages take longer than usual to open.",
              "Final status: A configuration change on part of the service caused the slowness. We reverted it and confirmed that page load times are back to normal.",
              "Scope of impact: A small number of users in this tenant's region were affected.",
            ]),
          },
        },
      ],
    },
  ];

  const messages = [
    {
      id: "MC900141",
      title: "Action required: decide how long anonymous sharing links should last",
      category: "planForChange",
      severity: "high",
      tags: ["Admin impact", "User impact"],
      services: ["SharePoint Online", "OneDrive for Business"],
      isMajorChange: false,
      actionRequiredByDateTime: iso(t + 10 * DAY),
      startDateTime: iso(t - 9 * DAY),
      endDateTime: null,
      lastModifiedDateTime: iso(t - 3 * DAY),
      body: {
        contentType: "html",
        content: html([
          "Microsoft is changing the default expiration for anonymous sharing links in SharePoint Online and OneDrive. Links created before the change keep working until their new expiration date unless an administrator sets a policy of their own.",
          "How this will affect your organization: Anonymous links will stop working sooner than they do today when no policy is set.",
          "What you need to do to prepare: Review the sharing settings in the SharePoint admin center, decide whether to keep the new default or set your own, and tell the teams that rely on long-lived links.",
        ]),
      },
    },
    {
      id: "MC900188",
      title: "Retirement: basic authentication for SMTP client submission in Exchange Online",
      category: "planForChange",
      severity: "high",
      tags: ["Retirement", "Admin impact"],
      services: ["Exchange Online"],
      isMajorChange: false,
      actionRequiredByDateTime: null,
      startDateTime: iso(t - 14 * DAY),
      endDateTime: null,
      lastModifiedDateTime: iso(t - 12 * DAY),
      body: {
        contentType: "html",
        content: html([
          "Basic authentication for SMTP client submission is being retired. Devices and applications that still send mail with a username and password will need to move to modern authentication.",
          "How this will affect your organization: Scanners, instruments and older applications that submit mail directly may stop sending.",
          "What you need to do to prepare: List the devices that submit mail through your tenant, check which of them support modern authentication, and plan a replacement for the rest.",
        ]),
      },
    },
    {
      id: "MC900202",
      title: "A redesigned meeting recap is coming to Microsoft Teams",
      category: "planForChange",
      severity: "normal",
      tags: ["New feature", "User impact"],
      services: ["Microsoft Teams"],
      isMajorChange: true,
      actionRequiredByDateTime: null,
      startDateTime: iso(t - 6 * DAY),
      endDateTime: null,
      lastModifiedDateTime: iso(t - 5 * DAY),
      body: {
        contentType: "html",
        content: html([
          "Meeting recaps in Microsoft Teams are getting a new layout that puts decisions and follow-up items first. The change rolls out to all users without any administrator action.",
          "How this will affect your organization: People will see the new layout the next time they open a recap.",
          "What you need to do to prepare: Nothing is required. You may want to tell your help desk about the new layout.",
        ]),
      },
    },
  ];
  return { overviews, issues, messages };
}

// ---------------------------------------------------------------------------------------------
// The closure check
// ---------------------------------------------------------------------------------------------

const documentationRanges = ["192.0.2.", "198.51.100.", "203.0.113."];

// The recording is committed to a public repository, so it is checked rather than trusted:
// every email address, link host and IP address in it must belong to the invented tenant, to
// the reserved .example suffix, to a documentation range, or be one of Microsoft's documented
// admin and documentation hosts. `known` is the set of addresses the tenant defines.
export function assertSynthetic(text: string, known: ReadonlySet<string>): void {
  const problems: string[] = [];
  for (const match of text.matchAll(/[A-Za-z0-9._%+#-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g)) {
    const address = match[0].toLowerCase();
    const domain = address.slice(address.lastIndexOf("@") + 1);
    if (!domain.endsWith(".example") || (domain === DOMAIN && !known.has(address)))
      problems.push(`address ${match[0]}`);
  }
  for (const match of text.matchAll(/https?:\/\/([^\s"\\/]+)/g)) {
    const host = (match[1] ?? "").toLowerCase();
    if (!host.endsWith(".example") && !(DOCUMENTED_HOSTS as readonly string[]).includes(host))
      problems.push(`link host ${host}`);
  }
  for (const match of text.matchAll(/\b(?:\d{1,3}\.){3}\d{1,3}\b/g))
    if (!documentationRanges.some((range) => match[0].startsWith(range)))
      problems.push(`IPv4 ${match[0]}`);
  for (const match of text.matchAll(
    /\b(?:[0-9a-f]{1,4}:){3,7}[0-9a-f]{1,4}\b|\b(?:[0-9a-f]{1,4}:){1,7}:(?:[0-9a-f]{1,4})?/gi,
  ))
    if (!match[0].toLowerCase().startsWith("2001:db8:")) problems.push(`IPv6 ${match[0]}`);
  if (problems.length > 0)
    throw new Error(
      `The recording contains text from outside the synthetic tenant: ${[...new Set(problems)]
        .slice(0, 10)
        .join(", ")}`,
    );
}

// ---------------------------------------------------------------------------------------------
// The Graph endpoint
// ---------------------------------------------------------------------------------------------

const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
const notFound = (message = "Resource not found for the segment requested.") =>
  json({ error: { code: "Request_ResourceNotFound", message } }, 404);

function project<T extends object>(item: T, select: string | null): Partial<T> {
  if (!select) return item;
  const wanted = new Set(["id", ...select.split(",").map((field) => field.trim())]);
  return Object.fromEntries(Object.entries(item).filter(([key]) => wanted.has(key))) as Partial<T>;
}

// One page of a collection with a $skiptoken continuation.
function page<T extends object>(url: URL, items: readonly T[], select: string | null) {
  const requested = Number(url.searchParams.get("$top") ?? "100");
  const top = Math.min(Math.max(Number.isFinite(requested) ? requested : 100, 1), 1_000);
  const offset = Number(url.searchParams.get("$skiptoken") ?? "0") || 0;
  const slice = items.slice(offset, offset + top);
  const body: { value: Partial<T>[]; "@odata.nextLink"?: string } = {
    value: slice.map((item) => project(item, select)),
  };
  if (offset + top < items.length) {
    const next = new URL(url);
    next.searchParams.set("$skiptoken", String(offset + top));
    body["@odata.nextLink"] = next.href;
  }
  return json(body);
}

function signInFilter(filter: string) {
  const clause = (pattern: RegExp) => pattern.exec(filter)?.[1];
  const from = clause(/createdDateTime ge ([0-9T:.\-Z+]+)/);
  const upper = /createdDateTime (le|lt) ([0-9T:.\-Z+]+)/.exec(filter);
  const user = clause(/userId eq '((?:[^']|'')*)'/)?.replaceAll("''", "'");
  const lowerBound = from ? Date.parse(from) : Number.NEGATIVE_INFINITY;
  const upperBound = upper?.[2] ? Date.parse(upper[2]) : Number.POSITIVE_INFINITY;
  const inclusive = upper?.[1] === "le";
  return (event: SignIn) => {
    const at = Date.parse(event.createdDateTime);
    return (
      at >= lowerBound &&
      (inclusive ? at <= upperBound : at < upperBound) &&
      (user === undefined || event.userId === user)
    );
  };
}

function issueFilter(filter: string | null) {
  if (!filter) return () => true;
  if (/isResolved eq false/i.test(filter))
    return (issue: { isResolved: boolean }) => !issue.isResolved;
  const since = /lastModifiedDateTime ge ([0-9T:.\-Z+]+)/.exec(filter)?.[1];
  if (since) {
    const bound = Date.parse(since);
    return (issue: { lastModifiedDateTime: string }) =>
      Date.parse(issue.lastModifiedDateTime) >= bound;
  }
  return () => true;
}

export function createSyntheticGraph(options: { now: Date }): typeof fetch {
  const tenant = buildSyntheticTenant(options.now);
  const communications = buildCommunications(options.now.getTime());
  const sortedSignIns = [...tenant.signIns].sort(
    (a, b) => b.createdDateTime.localeCompare(a.createdDateTime) || b.id.localeCompare(a.id),
  );
  const byId = new Map(tenant.users.map((user) => [user.id, user]));
  const byPrincipal = new Map(
    tenant.users.map((user) => [user.userPrincipalName.toLowerCase(), user]),
  );
  const licenseDetails = (user: DirectoryUser) =>
    user.assignedLicenses.map((license) => {
      const sku = tenant.skus.find((candidate) => candidate.skuId === license.skuId);
      return {
        id: guid(`license-detail:${user.id}:${license.skuId}`),
        skuId: license.skuId,
        skuPartNumber: sku?.skuPartNumber ?? license.skuId,
        servicePlans: (sku?.servicePlans ?? []).map(
          ({ servicePlanId, servicePlanName, provisioningStatus }) => ({
            servicePlanId,
            servicePlanName,
            provisioningStatus,
          }),
        ),
      };
    });

  return (async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    if (url.hostname !== "graph.microsoft.com" || !url.pathname.startsWith("/v1.0/"))
      return notFound();
    const path = decodeURIComponent(url.pathname.slice("/v1.0/".length)).replace(/\/$/, "");
    const select = url.searchParams.get("$select");

    if (path === "auditLogs/signIns") {
      const matches = signInFilter(url.searchParams.get("$filter") ?? "");
      return page(url, sortedSignIns.filter(matches), null);
    }
    if (path === "users") return page(url, tenant.users, select);
    const principal = /^users\('((?:[^']|'')*)'\)(\/.*)?$/.exec(path);
    const byPath = /^users\/([^/]+)(\/.*)?$/.exec(path);
    if (principal || byPath) {
      const user = principal
        ? byPrincipal.get((principal[1] ?? "").replaceAll("''", "'").toLowerCase())
        : byId.get(byPath?.[1] ?? "");
      const rest = (principal ? principal[2] : byPath?.[2]) ?? "";
      if (!user) return notFound();
      if (rest === "") return json(project(user, select));
      if (rest === "/manager") {
        const manager = byId.get(tenant.managers.get(user.id) ?? "");
        return manager
          ? json(project({ "@odata.type": "#microsoft.graph.user", ...manager }, select))
          : notFound();
      }
      if (rest === "/licenseDetails") return page(url, licenseDetails(user), select);
      return notFound();
    }
    if (path === "subscribedSkus") return page(url, tenant.skus, null);
    if (path === "reports/authenticationMethods/userRegistrationDetails")
      return page(url, tenant.registrations, null);
    if (path === "admin/serviceAnnouncement/healthOverviews")
      return page(url, communications.overviews, null);
    if (path === "admin/serviceAnnouncement/issues")
      return page(
        url,
        communications.issues.filter(issueFilter(url.searchParams.get("$filter"))),
        null,
      );
    if (path === "admin/serviceAnnouncement/messages")
      return page(url, communications.messages, null);
    return notFound();
  }) as typeof fetch;
}
