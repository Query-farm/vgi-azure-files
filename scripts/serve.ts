// Serve the vgi-azure-files worker over HTTP with the standardized VGI landing surface.
//
//   GET  /                                     → the shared vendored VGI landing.html
//   GET  /describe.json                        → the worker's catalog introspection
//   GET  /describe/{catalog}/{schema}/{t}.json → lazy per-object columns
//   GET  /health                               → JSON health endpoint (no credentials needed)
//   POST /                                     → the VGI RPC transport (what DuckDB attaches to)
//
// Run it:  PORT=8000 bun run scripts/serve.ts   (default port 8787)
// Attach:  ATTACH 'files' AS files (TYPE vgi, LOCATION 'http://localhost:8000',
//                 TENANT_ID '…', CLIENT_ID '…', CLIENT_SECRET '…');
//
// The content functions need an app-only azure_graph credential (tenant/client/
// secret, supplied as ATTACH options / a CREATE SECRET) to return rows — but the
// image itself just runs: GET /health and catalog introspection need no credentials.
//
// Everything below the worker's own identity — protocol assembly, state-token
// signing, CORS, the landing surface, Bun.serve — lives in the SDK's
// serveVgiWorker. Set VGI_SIGNING_KEY (64 hex chars) for any real deployment;
// without it the SDK generates an ephemeral key and warns.
//
// The wiring here mirrors src/worker.ts (the stdio entry): same MSAL-backed Graph
// client factory injected into the same delta table functions, same registry +
// catalog. Adding a function means updating BOTH entries.

import { serveVgiWorker } from "@query-farm/vgi/serve";
import { ReadOnlyCatalogInterface, FunctionRegistry } from "@query-farm/vgi";
import { TokenCache, makeGraphClient, type Fetch } from "@vgi-azure/graph-core";
import { makeMsalMinter } from "@vgi-azure/node-auth";
import { makeFunctions } from "../src/functions.js";
import { makeCatalog } from "../src/catalog.js";

const cache = new TokenCache(makeMsalMinter());

const clientFactory = (secret: Record<string, unknown>) =>
  makeGraphClient({
    fetch: globalThis.fetch as unknown as Fetch,
    cache,
    cred: {
      tenantId: String(secret.tenant_id ?? ""),
      clientId: String(secret.client_id ?? ""),
      clientSecret: secret.client_secret != null ? String(secret.client_secret) : undefined,
    },
    audience: "graph", // Microsoft Graph content endpoints
  });

const functions = makeFunctions(clientFactory);

const registry = new FunctionRegistry();
for (const f of functions) registry.register(f);

const catalogInterface = new ReadOnlyCatalogInterface(makeCatalog(functions), registry);

serveVgiWorker({
  name: "azure",
  doc: "Microsoft 365 content — OneDrive / SharePoint driveItems and SharePoint listItems — as incremental delta table functions over Microsoft Graph.",
  version: "0.1.0",
  repositoryUrl: "https://github.com/Query-farm/vgi-azure-files",
  serverId: "vgi-azure-files",
  registry,
  catalogInterface,
});
