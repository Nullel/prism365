import { capabilitySummarySchema } from "@prism365/contracts";
import { microsoftCloudProfiles, permissionPacks } from "@prism365/graph";
import Fastify from "fastify";
import { demoDashboard } from "./demo-data.ts";

export function createApp() {
  const app = Fastify({
    logger: false,
    requestIdHeader: "x-request-id",
  });

  app.addHook("onSend", async (_request, reply, payload) => {
    reply.header("Cache-Control", "no-store");
    reply.header("X-Content-Type-Options", "nosniff");
    reply.header("X-Frame-Options", "DENY");
    reply.header("Referrer-Policy", "no-referrer");
    return payload;
  });

  app.get("/health", async () => ({
    status: "ok",
    mode: "demo",
    tenantConnected: false,
  }));

  app.get("/v1/dashboard/summary", async () => demoDashboard);

  app.get("/v1/platform/capabilities", async () =>
    capabilitySummarySchema.parse({
      dataMode: "demo",
      writesEnabled: false,
      tenantConnectionEnabled: false,
      permissionPacks,
      cloudProfiles: Object.values(microsoftCloudProfiles),
      ai: {
        analysisEnabled: false,
        providerConfigured: false,
        writeToolsEnabled: false,
        supportedDeployments: ["local", "tenantHosted", "external"],
      },
    }),
  );

  return app;
}
