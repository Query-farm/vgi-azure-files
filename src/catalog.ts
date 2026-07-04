// The `azure` catalog descriptor + the azure_graph secret type for vgi-azure-files.
// The secret shape is the frozen graph-core seam (app-only client-credentials) that
// vgi-azure-directory owns; files reuses it verbatim so every worker shares one secret
// type across the catalog.

import { Schema, Field, Utf8 } from "@query-farm/apache-arrow";
import type { CatalogDescriptor, SecretTypeDescriptor, VgiFunction } from "@query-farm/vgi";

export const AZURE_GRAPH_SECRET: SecretTypeDescriptor = {
  name: "azure_graph",
  description: "Microsoft Entra app-only (client-credentials) credentials for Microsoft Graph",
  schema: new Schema([
    new Field("tenant_id", new Utf8(), true),
    new Field("client_id", new Utf8(), true),
    new Field("client_secret", new Utf8(), true, new Map([["redact", "true"]])),
  ]),
};

export function makeCatalog(functions: VgiFunction[]): CatalogDescriptor {
  return {
    name: "azure",
    defaultSchema: "main",
    // v1 least-privilege posture (SPEC §3a, SECURITY): ship Sites.Selected — the app is
    // granted read on only admin-allow-listed sites (per-site grant via
    // /sites/{id}/permissions). Tenant-wide Sites.Read.All / Files.Read.All is opt-in for
    // fleet-wide scans, not the default. _download_url is opt-in-only (SPEC §4a).
    comment:
      "Microsoft 365 content (OneDrive / SharePoint driveItems + listItems) via Graph per-resource " +
      "delta — vgi-azure-files. v1 least-privilege: Sites.Selected (opt-in Sites.Read.All for fleet scans).",
    sourceUrl: "https://query.farm",
    secretTypes: [AZURE_GRAPH_SECRET],
    schemas: [{ name: "main", functions }],
  };
}
