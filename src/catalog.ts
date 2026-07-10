// The `azure` catalog descriptor + the azure_graph secret type for vgi-azure-files.
// The secret shape is the frozen graph-core seam (app-only client-credentials) that
// vgi-azure-directory owns; files reuses it verbatim so every worker shares one secret
// type across the catalog.
//
// This file also carries the catalog- and schema-level vgi.* documentation/discovery
// tags that vgi-lint grades. Tag shapes follow vgi-lint's TAGS.md: JSON-valued tags
// (keywords/categories/executable_examples/agent_test_tasks/example_queries) are JSON
// strings; every example SQL is catalog-qualified (azure.main.<fn>) so it binds when the
// catalog is attached. The content functions require an `azure_graph` secret + a network
// call to RUN, so the executable examples are credential-free `LIMIT 0` bind probes and
// the agent-test tasks are graded by success_criteria (credential-gated / non-deterministic).

import { Schema, Field, Utf8 } from "@query-farm/apache-arrow";
import type { CatalogDescriptor, SecretTypeDescriptor, ViewDescriptor, VgiFunction } from "@query-farm/vgi";

const REPO = "https://github.com/Query-farm/vgi-azure-files";
const ISSUES = `${REPO}/issues`;

export const AZURE_GRAPH_SECRET: SecretTypeDescriptor = {
  name: "azure_graph",
  description: "Microsoft Entra app-only (client-credentials) credentials for Microsoft Graph",
  schema: new Schema([
    new Field("tenant_id", new Utf8(), true),
    new Field("client_id", new Utf8(), true),
    new Field("client_secret", new Utf8(), true, new Map([["redact", "true"]])),
  ]),
};

/** Catalog-level tags: docs, discovery, provenance, and the agent-test suite. */
const CATALOG_TAGS: Record<string, string> = {
  "vgi.title": "Microsoft 365 Content Delta",
  "vgi.doc_llm":
    "Microsoft 365 file/content objects as SQL delta table functions over Microsoft Graph. Reach for " +
    "it to sync a SharePoint / OneDrive document library or a SharePoint list as an incremental change " +
    "feed: each scan returns the items created/updated (change_type 'upsert') or deleted (change_type " +
    "'removed') since the last cursor, plus a marker row carrying the next delta cursor to persist. " +
    "Each function is a per-resource cursor atom — you pass ONE delta_token for ONE drive/list and get " +
    "ONE _delta_next back; the caller holds the {resourceId -> deltaLink} map and fans out across the " +
    "fleet. A first call with no delta_token performs a full sync; passing a previously saved " +
    "delta_token returns only what changed. Requires an app-only (client-credentials) 'azure_graph' " +
    "secret (tenant_id, client_id, client_secret) with Graph read permission for the resources being " +
    "queried (e.g. Sites.Selected, Sites.Read.All, or Files.Read.All).",
  "vgi.doc_md":
    "## Microsoft 365 Content Delta\n\n" +
    "Incremental (delta) access to Microsoft 365 content via Microsoft Graph, exposed as two DuckDB " +
    "table functions.\n\n" +
    "- **`drive_items`** — OneDrive / SharePoint document-library `driveItems` as a per-drive " +
    "incremental change feed (with removals); optional opt-in `_download_url` capability column.\n" +
    "- **`sharepoint_list_items`** — SharePoint `listItems` as a per-list incremental change feed (with " +
    "removals); `$expand=fields` materializes user-defined columns.\n\n" +
    "Each function returns changed items plus a single marker row (`_row_kind = 'marker'`) whose " +
    "`_delta_next` column holds the verbatim `@odata.deltaLink` to persist and replay on the next " +
    "scan. Call with no `delta_token` for a full sync; pass `delta_token := '<link>'` for an " +
    "incremental sync. Each invocation is a single-resource cursor atom — the caller holds the " +
    "`{resourceId -> deltaLink}` map. An app-only `azure_graph` secret (Microsoft Entra client " +
    "credentials) is required.",
  "vgi.keywords": JSON.stringify([
    "azure",
    "microsoft 365",
    "microsoft graph",
    "sharepoint",
    "onedrive",
    "drive items",
    "list items",
    "files",
    "documents",
    "content",
    "delta",
    "change feed",
    "cdc",
  ]),
  "vgi.author": "Query Farm LLC",
  "vgi.copyright": "Copyright 2026 Query Farm LLC",
  "vgi.license": "MIT",
  "vgi.support_contact": ISSUES,
  "vgi.support_policy_url": ISSUES,
  // Guaranteed-runnable, catalog-qualified examples (VGI509/VGI906). A LIVE sync needs an
  // attached azure_graph secret and a network call to Microsoft Graph, so these are
  // credential-free `LIMIT 0` bind probes: onBind runs (the required args are supplied so
  // it succeeds) and exposes each function's exact result columns, but process() — where
  // the secret + network live — is never pumped. Drop the `LIMIT 0`, supply real
  // drive/site/list ids, and attach an azure_graph secret to pull real rows — the fuller,
  // data-returning queries live in each function's `examples` and the schema
  // `vgi.example_queries`.
  "vgi.executable_examples": JSON.stringify([
    {
      name: "drive_items_bind_probe",
      description:
        "Bind drive_items and expose its default result columns (credential-free; the 'drive' placeholder satisfies onBind, LIMIT 0 skips the network. Supply a real driveId and attach an azure_graph secret to sync).",
      sql: "SELECT id, name, size, web_url, last_modified_date_time FROM azure.main.drive_items('placeholder-drive-id') LIMIT 0",
    },
    {
      name: "drive_items_download_url_bind_probe",
      description:
        "Bind drive_items with include_download_url := true to expose the opt-in _download_url capability column (credential-free; LIMIT 0 skips the network).",
      sql: "SELECT id, name, _download_url FROM azure.main.drive_items('placeholder-drive-id', include_download_url := true) LIMIT 0",
    },
    {
      name: "sharepoint_list_items_bind_probe",
      description:
        "Bind sharepoint_list_items and expose its result columns (credential-free; the 'site'/'list' placeholders satisfy onBind, LIMIT 0 skips the network. Supply real siteId/listId and attach an azure_graph secret to sync).",
      sql: "SELECT id, name, web_url, last_modified_date_time FROM azure.main.sharepoint_list_items('placeholder-site-id', 'placeholder-list-id') LIMIT 0",
    },
  ]),
  // The agent-suitability suite (VGI152/VGI520/VGI920), catalog only. Every worker object
  // (both delta functions + the browsable content_collections view) is exercised by exactly
  // one task. `reference_sql` is the canonical fully-qualified solution (grader-only): it
  // drives static coverage (VGI520) and, when the sim runs authenticated, reference grading.
  // `success_criteria` is the LLM-judge rubric used when the sim runs unauthenticated — a
  // live content scan needs an azure_graph secret and returns tenant-specific, non-
  // deterministic data, so an exact-compare reference can only be graded with real
  // credentials. Neither reference_sql nor success_criteria is ever shown to the analyst.
  "vgi.agent_test_tasks": JSON.stringify([
    {
      name: "list_drive_files",
      prompt:
        "List the files in the OneDrive / SharePoint document library with drive id '<driveId>', showing each file's name and size.",
      reference_sql:
        "SELECT name, size FROM azure.main.drive_items('<driveId>') WHERE _row_kind IS NULL ORDER BY name",
      success_criteria:
        "The answer queries drive_items('<driveId>') (the required drive id is passed positionally), filters to data rows (_row_kind IS NULL), and returns name and size (bytes).",
    },
    {
      name: "save_drive_cursor",
      prompt: "After scanning a drive with drive_items, how do I get the cursor to use for the next incremental sync?",
      reference_sql:
        "SELECT _delta_next FROM azure.main.drive_items('<driveId>') WHERE _row_kind = 'marker'",
      success_criteria:
        "The answer selects _delta_next from the marker row (_row_kind = 'marker') of drive_items() and explains it should be persisted per drive and replayed via the delta_token argument.",
    },
    {
      name: "incremental_list_changes",
      prompt:
        "Using a previously saved delta cursor, fetch only the SharePoint list items in site '<siteId>' list '<listId>' that changed since the last scan, including which ones were removed.",
      reference_sql:
        "SELECT id, change_type, name FROM azure.main.sharepoint_list_items('<siteId>', '<listId>', delta_token := '<@odata.deltaLink>')",
      success_criteria:
        "The answer queries sharepoint_list_items('<siteId>', '<listId>', delta_token := '<link>') (site and list ids passed positionally), filters to data rows (_row_kind IS NULL), and distinguishes change_type 'upsert' from 'removed'.",
    },
    {
      name: "browse_content_collections",
      prompt: "What kinds of Microsoft 365 content can I sync with this catalog, and which table function serves each?",
      reference_sql:
        "SELECT collection, table_function FROM azure.main.content_collections ORDER BY collection",
      success_criteria:
        "The answer reads content_collections (the browsable, credential-free discovery view) and lists the content kinds (document-library drive items, SharePoint list items) alongside the table function that serves each.",
    },
  ]),
};

/** Schema-level tags: docs, discovery, the category registry, and shown examples. */
const SCHEMA_TAGS: Record<string, string> = {
  "vgi.title": "Microsoft 365 Content",
  "vgi.doc_llm":
    "The Microsoft 365 content delta functions. Each function is an incremental change feed for one " +
    "kind of resource — a document library's driveItems, or a SharePoint list's listItems — driven by " +
    "Microsoft Graph's delta query. A scan returns rows for items created/updated (change_type " +
    "'upsert') or deleted (change_type 'removed') since the supplied delta_token, followed by one " +
    "marker row whose _delta_next column is the cursor to persist for the next scan. Omit delta_token " +
    "for a full sync; the $select / $expand / $top projection is pinned for the life of a token. Each " +
    "call is a single-resource cursor atom — the {resourceId -> deltaLink} map is held by the caller.",
  "vgi.doc_md":
    "## Microsoft 365 content delta functions\n\n" +
    "| Function | Resource | Returns |\n" +
    "| --- | --- | --- |\n" +
    "| `drive_items` | OneDrive / SharePoint document library | changed driveItems + delta cursor |\n" +
    "| `sharepoint_list_items` | SharePoint list | changed listItems + delta cursor |\n\n" +
    "Both share the same shape: read data rows with `WHERE _row_kind IS NULL`, and take the next " +
    "cursor from the single marker row's `_delta_next`. `drive_items` can additionally opt into a " +
    "short-lived `_download_url` capability column via `include_download_url := true`. Each requires " +
    "an app-only `azure_graph` secret.",
  "vgi.keywords": JSON.stringify([
    "microsoft 365",
    "microsoft graph",
    "sharepoint",
    "onedrive",
    "drive items",
    "list items",
    "files",
    "documents",
    "content",
    "delta query",
    "change feed",
  ]),
  domain: "content",
  // Ordered navigation registry; each `name` is referenced by an object's vgi.category.
  "vgi.categories": JSON.stringify([
    {
      name: "discovery",
      title: "Discovery",
      description:
        "Browsable, credential-free entry points for finding your way around the content catalog before attaching a secret.",
    },
    {
      name: "drive-items",
      title: "Document Library Delta",
      description:
        "Incremental change feed of OneDrive / SharePoint document-library driveItems (files and folders) via Microsoft Graph delta query.",
    },
    {
      name: "list-items",
      title: "SharePoint List Delta",
      description:
        "Incremental change feed of SharePoint list items (with user-defined columns via $expand=fields) via Microsoft Graph delta query.",
    },
  ]),
  "vgi.example_queries": JSON.stringify([
    {
      description: "Full sync of a document library's files (data rows only)",
      sql: "SELECT id, name, size, web_url FROM azure.main.drive_items('<driveId>') WHERE _row_kind IS NULL",
    },
    {
      description: "Incremental sync of a drive from a saved cursor (upserts and removals)",
      sql: "SELECT id, change_type, name FROM azure.main.drive_items('<driveId>', delta_token := '<@odata.deltaLink>')",
    },
    {
      description: "Include the opt-in short-lived download capability URL",
      sql: "SELECT id, name, _download_url FROM azure.main.drive_items('<driveId>', include_download_url := true) WHERE _row_kind IS NULL AND change_type = 'upsert'",
    },
    {
      description: "Full sync of a SharePoint list's items",
      sql: "SELECT id, name, web_url, last_modified_date_time FROM azure.main.sharepoint_list_items('<siteId>', '<listId>') WHERE _row_kind IS NULL",
    },
    {
      description: "Read the delta cursor to persist for a drive's next sync",
      sql: "SELECT _delta_next FROM azure.main.drive_items('<driveId>') WHERE _row_kind = 'marker'",
    },
  ]),
};

// A browsable, credential-free discovery view: the content collections this catalog
// exposes, the delta table function that serves each, and a one-line description. Its
// definition is a self-contained VALUES relation evaluated entirely by DuckDB (no worker
// call, no secret), so an agent can `SELECT * FROM azure.main.content_collections` to learn
// the surface before it ever needs Microsoft Graph credentials. This is the worker's
// browsable entry point (VGI146): every other object here is a credential-gated table
// function.
const CONTENT_COLLECTIONS_VIEW: ViewDescriptor = {
  name: "content_collections",
  definition:
    "SELECT collection, resource, table_function, description FROM (VALUES " +
    "('drive_items', 'driveItem', 'drive_items', 'OneDrive / SharePoint document-library files and folders as an incremental Microsoft Graph delta change feed'), " +
    "('list_items', 'listItem', 'sharepoint_list_items', 'SharePoint list items (with user-defined columns via $expand=fields) as an incremental Microsoft Graph delta change feed')" +
    ") AS t(collection, resource, table_function, description)",
  comment:
    "The Microsoft 365 content collections this catalog exposes (document-library drive items, SharePoint list items) and the delta table function that serves each. Browsable without credentials.",
  columnComments: {
    collection: "The collection slug (drive_items / list_items).",
    resource: "The singular Microsoft Graph resource type the collection contains (driveItem / listItem).",
    table_function: "The catalog table function that syncs this collection as a delta feed.",
    description: "A one-line description of the collection.",
  },
  tags: {
    "vgi.title": "Content Collection Index",
    "vgi.category": "discovery",
    domain: "content",
    "vgi.doc_llm":
      "A static, credential-free catalog of the Microsoft 365 content collections this worker exposes: one " +
      "row per collection (document-library drive items, SharePoint list items) giving its Graph resource " +
      "type, the table function that syncs it as an incremental Microsoft Graph delta feed, and a short " +
      "description. Query it to discover the worker's surface before attaching an azure_graph secret.",
    "vgi.doc_md":
      "## content_collections\n\n" +
      "A browsable, credential-free index of the Microsoft 365 content collections this catalog exposes. " +
      "One row per collection, naming the delta table function that syncs it. Start here, then call the " +
      "named function (with an `azure_graph` secret attached) to sync that collection.",
    "vgi.keywords": JSON.stringify([
      "content",
      "collections",
      "catalog",
      "discovery",
      "drive items",
      "list items",
      "table functions",
    ]),
    "vgi.example_queries": JSON.stringify([
      {
        description: "List every content collection and the table function that serves it",
        sql: "SELECT collection, table_function FROM azure.main.content_collections ORDER BY collection",
      },
      {
        description: "Find the table function for a given Graph resource type",
        sql: "SELECT table_function FROM azure.main.content_collections WHERE resource = 'listItem'",
      },
    ]),
  },
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
    sourceUrl: REPO,
    tags: CATALOG_TAGS,
    secretTypes: [AZURE_GRAPH_SECRET],
    schemas: [
      {
        name: "main",
        comment: "Microsoft 365 content (OneDrive / SharePoint driveItems + SharePoint listItems) as incremental Graph delta feeds.",
        tags: SCHEMA_TAGS,
        views: [CONTENT_COLLECTIONS_VIEW],
        functions,
      },
    ],
  };
}
