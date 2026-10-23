import { type AuditRow, licenseName } from "@prism365/contracts";
import type { Audit } from "./definition.ts";
import { readSkus, tagOf } from "./directory.ts";

// Products Microsoft provisions with this many seats or more are free or self-service sign-up
// products (Power Automate Free, Fabric Free); their unused seats are not worth reporting.
const FREE_PRODUCT_SEATS = 10_000;

const statusLabels: Record<string, string> = {
  enabled: "Active",
  warning: "Expired, in grace period",
  suspended: "Suspended",
  lockedout: "Locked out",
  deleted: "Deleted",
};

// Each product the tenant subscribes to: seats bought, assigned and free.
export const licenseCapacityAudit: Audit = {
  definition: {
    id: "licenses.capacity",
    version: 1,
    title: "License capacity",
    description:
      "Each product the tenant subscribes to: seats bought, assigned and free, and subscriptions that have expired or been suspended.",
    category: "Licenses",
    subject: { one: "product", many: "products" },
    columns: [
      { key: "product", label: "Product", kind: "text" },
      { key: "part", label: "Part number", kind: "text" },
      { key: "status", label: "Subscription", kind: "enum", facet: true },
      { key: "active", label: "Active seats", kind: "number" },
      {
        key: "grace",
        label: "In grace period",
        kind: "number",
        help: "Seats on an expired subscription that can still be assigned until it is renewed or suspended.",
      },
      { key: "assigned", label: "Assigned", kind: "number" },
      {
        key: "available",
        label: "Available",
        kind: "number",
        help: "Active and grace-period seats not assigned to anyone. Negative means more are assigned than owned.",
      },
      { key: "suspended", label: "Suspended", kind: "number", hidden: true },
      { key: "lockedOut", label: "Locked out", kind: "number", hidden: true },
      { key: "appliesTo", label: "Applies to", kind: "enum", facet: true, hidden: true },
    ],
    views: [
      { id: "attention", label: "Needs attention", tags: ["attention"] },
      { id: "all", label: "All products", tags: [] },
      { id: "full", label: "No seats left", tags: ["full"] },
      { id: "spare", label: "Unassigned seats", tags: ["spare"] },
      { id: "expiring", label: "Expired or suspended", tags: ["expiring"] },
    ],
    defaultView: "attention",
    defaultSort: { column: "available", direction: "asc" },
    requirements: [
      "A role that can read subscriptions, such as Directory Readers, Global Reader or License Administrator.",
    ],
  },
  packs: ["tenant.licenses.read"],
  optionalPacks: [],
  scoped: false,
  forbidden:
    "Microsoft refused to list subscriptions. It needs LicenseAssignment.Read.All and a role such as Directory Readers, Global Reader or License Administrator.",
  async run(context) {
    const read = await readSkus(context);
    const rows: AuditRow[] = read.skus.map((sku) => {
      const units = sku.prepaidUnits;
      const owned = units.enabled + units.warning;
      const available = owned - sku.consumedUnits;
      const perUser = (sku.appliesTo ?? "user").toLowerCase() === "user";
      const status = sku.capabilityStatus ?? "Unknown";
      const expiring =
        units.warning > 0 ||
        units.suspended > 0 ||
        units.lockedOut > 0 ||
        status.toLowerCase() !== "enabled";
      const full = perUser && owned > 0 && available <= 0;
      // Seats on an expired or suspended subscription are not spare: they are going away.
      const spare = perUser && available > 0 && owned < FREE_PRODUCT_SEATS && !expiring;
      return {
        key: sku.skuId,
        personId: null,
        tags: [
          ...(full || expiring || available < 0 ? ["attention"] : []),
          ...(full ? ["full"] : []),
          ...(spare ? ["spare"] : []),
          ...(expiring ? ["expiring"] : []),
          ...tagOf(sku.appliesTo),
        ],
        cells: {
          product: licenseName(sku.skuPartNumber ?? sku.skuId),
          part: sku.skuPartNumber,
          status: statusLabels[status.toLowerCase()] ?? status,
          active: units.enabled,
          grace: units.warning,
          assigned: sku.consumedUnits,
          available,
          suspended: units.suspended,
          lockedOut: units.lockedOut,
          appliesTo: sku.appliesTo,
        },
      };
    });
    rows.sort((a, b) => String(a.cells.product).localeCompare(String(b.cells.product)));
    const count = (tag: string) => rows.filter((row) => row.tags.includes(tag)).length;
    const spareSeats = rows
      .filter((row) => row.tags.includes("spare"))
      .reduce((total, row) => total + Number(row.cells.available ?? 0), 0);
    const attention = count("attention");
    return {
      rows,
      complete: read.complete && read.skipped === 0,
      summary: [
        { label: "Products", value: rows.length, tone: "neutral", view: "all" },
        {
          label: "Needs attention",
          value: attention,
          tone: attention > 0 ? "critical" : "good",
          view: "attention",
        },
        { label: "No seats left", value: count("full"), tone: "warning", view: "full" },
        { label: "Unassigned seats", value: spareSeats, tone: "neutral", view: "spare" },
        {
          label: "Expired or suspended",
          value: count("expiring"),
          tone: "warning",
          view: "expiring",
        },
      ],
      notes: [
        "Needs attention: no seats left to assign, more assigned than owned, or a subscription that has expired, been suspended or been locked out.",
        `Products with ${FREE_PRODUCT_SEATS.toLocaleString("en-US")} seats or more are usually free or self-service sign-up products, so their seats are not counted as unassigned.`,
        "Microsoft does not report prices or renewal dates here, so Prism365 shows seats only.",
        ...(read.skipped > 0
          ? [
              `${read.skipped} ${read.skipped === 1 ? "product" : "products"} could not be read and are left out.`,
            ]
          : []),
      ],
    };
  },
};
