// vgi-azure-files stdio worker entry. DuckDB spawns this and ATTACHes it:
//   ATTACH 'files' AS files (TYPE vgi, LOCATION '/path/to/worker.ts');
//   CREATE SECRET g (TYPE azure_graph, TENANT_ID '…', CLIENT_ID '…', CLIENT_SECRET '…');
//   -- initial full sync of one drive (runs on Node: Graph emits the deltaLink only on
//   -- the FINAL page, so a first sync has NO mid-sync checkpoint — SPEC §6a EDGE):
//   SELECT * FROM files.drive_items(drive := '<driveId>') WHERE _row_kind IS NULL;
//   -- incremental (replay the persisted verbatim @odata.deltaLink for that drive):
//   SELECT * FROM files.drive_items(drive := '<driveId>', delta_token := '<link>');
//   SELECT * FROM files.list_items(site := '<siteId>', list := '<listId>');
// The caller holds the {resourceId -> deltaLink} MAP and feeds one token per call.

import { Worker, ReadOnlyCatalogInterface, FunctionRegistry } from "@query-farm/vgi";
import { TokenCache, makeGraphClient, type Fetch } from "@vgi-azure/graph-core";
import { makeMsalMinter } from "@vgi-azure/node-auth";
import { makeFunctions } from "./functions.js";
import { makeCatalog } from "./catalog.js";

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

new Worker({ functions, catalogInterface }).run();
