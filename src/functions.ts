// The VGI table functions: drive_items + sharepoint_list_items. Both are TABLE functions so
// `name := value` works (scalars are positional-only — SPEC §4 / graph-core checklist).
// The GraphClient is injected via a ClientFactory so the worker wires the real
// MSAL-backed client and tests inject a fake.
//
// Multi-cursor MAP: each function is a per-resource cursor atom — ONE bare delta_token
// in, ONE _delta_next out. The {resourceId -> deltaLink} map lives on the CALLER (§2b);
// the worker never iterates the fleet, so every invocation is independently resumable.

import { defineTableFunction, secretsOfType, type OutputCollector } from "@query-farm/vgi";
import { Utf8, Int64, Bool } from "@query-farm/apache-arrow";
import { collectDelta, driveItemsStartUrl, listItemsStartUrl } from "./files-delta.js";
import { fileItemSchema, buildDeltaBatch } from "./schema.js";
import type { GraphClient } from "@vgi-azure/graph-core";

export type ClientFactory = (secret: Record<string, unknown>) => GraphClient;

/** Best-effort $select on delta (tolerate extra fields; do NOT assume $top honored). */
const DEFAULT_ITEM_SELECT = "id,name,size,webUrl,lastModifiedDateTime";
const DEFAULT_EXPAND = "fields";
const DEFAULT_PAGE_SIZE = 200;

function requireSecret(p: { secrets: Record<string, Record<string, unknown>> }, fn: string): Record<string, unknown> {
  const secret = secretsOfType(p.secrets, "azure_graph")[0];
  if (!secret) throw new Error(`${fn}: attach an 'azure_graph' secret (TYPE azure_graph)`);
  return secret as Record<string, unknown>;
}

// Structured result-schema columns, kept in one place so both functions' declared
// schemas stay consistent with fileItemSchema() + buildDeltaBatch(). These feed the
// modern structured tags (VGI307): sharepoint_list_items has a static schema (vgi.result_columns_schema),
// while drive_items varies by include_download_url (vgi.result_dynamic_columns_md).
interface ResultColumn {
  name: string;
  type: string; // a real DuckDB type (VGI322)
  description: string;
}
const BUSINESS_RESULT_COLUMNS: ResultColumn[] = [
  { name: "id", type: "VARCHAR", description: "The item's id. NULL on the marker row." },
  { name: "change_type", type: "VARCHAR", description: "`upsert` (created/updated) or `removed` (deleted). NULL on the marker row." },
  { name: "removed_reason", type: "VARCHAR", description: "For removed rows, the `@removed.reason` (`changed` or `deleted`); NULL otherwise." },
  { name: "name", type: "VARCHAR", description: "The item's name (file/folder name, or the list item's Title). NULL on removed tombstones and the marker row." },
  { name: "size", type: "BIGINT", description: "The item's size in bytes. NULL for folders, removed tombstones, and the marker row." },
  { name: "web_url", type: "VARCHAR", description: "Browser URL to the item. NULL on removed tombstones and the marker row." },
  { name: "last_modified_date_time", type: "VARCHAR", description: "ISO 8601 last-modified timestamp. NULL on removed tombstones and the marker row." },
];
const DOWNLOAD_URL_RESULT_COLUMN: ResultColumn = {
  name: "_download_url",
  type: "VARCHAR",
  description:
    "Only present when `include_download_url := true`. A short-lived, pre-authed, no-bearer download capability URL for live file rows; NULL for folders, removed tombstones, and the marker row.",
};
const CONTROL_RESULT_COLUMNS: ResultColumn[] = [
  { name: "_row_kind", type: "VARCHAR", description: "NULL for data rows; `marker` for the single trailing cursor row." },
  {
    name: "_delta_next",
    type: "VARCHAR",
    description:
      "On the marker row, the verbatim `@odata.deltaLink` to persist and replay as this resource's next cursor; NULL on data rows.",
  },
];

/** Render a Name | Type | Description Markdown table (one variant of a dynamic schema). */
function columnsTable(cols: ResultColumn[]): string {
  return (
    "| Name | Type | Description |\n| --- | --- | --- |\n" +
    cols.map((c) => `| \`${c.name}\` | ${c.type} | ${c.description} |`).join("\n")
  );
}

// ---------------------------------------------------------------------------
// drive_items(drive, delta_token?, select?, page_size?, include_download_url?)
// ---------------------------------------------------------------------------

interface DriveArgs {
  drive: string;
  /** Full opaque @odata.deltaLink from last scan, or "" for a full sync. */
  delta_token: string;
  select: string;
  page_size: number; // Int64 args arrive coerced to number by the SDK
  include_download_url: boolean; // SECURITY: opt-in only (SPEC §4a)
}
interface DriveState {
  done: boolean;
  startUrl: string;
  resourceId: string;
  includeDownloadUrl: boolean;
}

export function makeDriveItems(clientFactory: ClientFactory) {
  return defineTableFunction<DriveArgs, DriveState>({
    name: "drive_items",
    description:
      "OneDrive / SharePoint document-library driveItems via Microsoft Graph per-drive delta " +
      "(incremental change feed with removals). One delta_token per drive; the {drive -> deltaLink} " +
      "map is held by the caller (multi-cursor). include_download_url is opt-in only (capability URL).",
    args: {
      drive: new Utf8(),
      delta_token: new Utf8(),
      select: new Utf8(),
      page_size: new Int64(),
      include_download_url: new Bool(),
    },
    argDefaults: {
      delta_token: "",
      select: DEFAULT_ITEM_SELECT,
      page_size: DEFAULT_PAGE_SIZE,
      include_download_url: false,
    },
    argDocs: {
      drive: "REQUIRED. The Microsoft Graph drive id of the OneDrive / SharePoint document library to scan (the `driveId`).",
      delta_token:
        "A previously persisted `@odata.deltaLink` (the `_delta_next` value from a prior scan's marker row for THIS drive), replayed verbatim to return only what changed since. Empty (the default) performs a full sync.",
      select:
        "Comma-separated Microsoft Graph `$select` projection of driveItem fields to fetch on the initial full sync. Pinned for the life of a delta token (baked into the server's deltaLink). Defaults to `id,name,size,webUrl,lastModifiedDateTime`.",
      page_size:
        "Best-effort Microsoft Graph `$top` page size for the initial full sync (Graph may not honor it). Defaults to 200.",
      include_download_url:
        "Opt-in only. When `true`, adds a `_download_url` column carrying a short-lived, pre-authed, no-bearer capability URL for each live file row (never folders or removed tombstones). Omitted from the schema entirely when `false` (the default) — SECURITY: this URL bypasses RBAC if leaked.",
    },
    examples: [
      {
        sql: "SELECT id, name, size, web_url FROM azure.main.drive_items('<driveId>') WHERE _row_kind IS NULL",
        description: "Full sync of a document library's files (data rows only)",
      },
      {
        sql: "SELECT id, change_type, name FROM azure.main.drive_items('<driveId>', delta_token := '<@odata.deltaLink>')",
        description: "Incremental sync replaying a previously saved delta cursor for this drive",
      },
      {
        sql: "SELECT id, name, _download_url FROM azure.main.drive_items('<driveId>', include_download_url := true) WHERE _row_kind IS NULL AND change_type = 'upsert'",
        description: "Fetch files with the opt-in short-lived download capability URL",
      },
      {
        sql: "SELECT _delta_next FROM azure.main.drive_items('<driveId>') WHERE _row_kind = 'marker'",
        description: "Read the delta cursor to persist for this drive's next sync",
      },
    ],
    tags: {
      "vgi.category": "drive-items",
      "vgi.title": "Document Library Delta Feed",
      "vgi.keywords": JSON.stringify([
        "onedrive",
        "sharepoint",
        "drive items",
        "driveitem",
        "document library",
        "files",
        "folders",
        "documents",
        "delta",
        "change feed",
        "download url",
      ]),
      "vgi.doc_llm":
        "Incremental change feed of OneDrive / SharePoint document-library driveItems (files and " +
        "folders) via Microsoft Graph per-drive delta query. Each scan returns items created or updated " +
        "since the supplied delta_token (change_type 'upsert') and items deleted (change_type " +
        "'removed'), followed by a marker row whose _delta_next is the cursor to persist for THIS drive. " +
        "Call with no delta_token for a full sync. `drive` is required (the Graph driveId). Columns: " +
        "name, size, web_url, last_modified_date_time; opt into a short-lived _download_url capability " +
        "column with include_download_url := true.",
      "vgi.doc_md":
        "## drive_items\n\n" +
        "OneDrive / SharePoint document-library `driveItems` as an incremental change feed backed by " +
        "Microsoft Graph's `drives/{drive}/root/delta` query. `drive` is required. Read data rows with " +
        "`WHERE _row_kind IS NULL`; take the next cursor from the marker row's `_delta_next` and persist " +
        "it per drive. Set `include_download_url := true` to add a short-lived download capability URL " +
        "column. See the examples for runnable full-sync, incremental, and download-URL queries.",
      // Dynamic: the column set varies by include_download_url, so one variant table per
      // shape (VGI307 dynamic branch / VGI326).
      "vgi.result_dynamic_columns_md":
        "The result schema depends on the `include_download_url` argument.\n\n" +
        "### Default (`include_download_url := false`)\n\n" +
        columnsTable([...BUSINESS_RESULT_COLUMNS, ...CONTROL_RESULT_COLUMNS]) +
        "\n\n### With `include_download_url := true`\n\n" +
        columnsTable([...BUSINESS_RESULT_COLUMNS, DOWNLOAD_URL_RESULT_COLUMN, ...CONTROL_RESULT_COLUMNS]),
    },
    onBind: (p) => {
      if (!p.args.drive) throw new Error("drive_items: 'drive' is required");
      // Schema is gated: _download_url present ONLY when opted in (SPEC §4a).
      return { outputSchema: fileItemSchema(Boolean(p.args.include_download_url)) };
    },
    initialState: (p) => ({
      done: false,
      resourceId: p.args.drive,
      includeDownloadUrl: Boolean(p.args.include_download_url),
      // delta_token, when present, IS the full opaque deltaLink -> replay VERBATIM.
      // $select/$top apply only to the initial-sync URL; on delta they are baked into
      // the server's link (SPEC §2a).
      startUrl: p.args.delta_token
        ? p.args.delta_token
        : driveItemsStartUrl(p.args.drive, p.args.select, Number(p.args.page_size) || DEFAULT_PAGE_SIZE),
    }),
    process: async (p, state: DriveState, out: OutputCollector) => {
      if (state.done) {
        out.finish();
        return;
      }
      const client = clientFactory(requireSecret(p, "drive_items"));
      const schema = fileItemSchema(state.includeDownloadUrl);
      const { rows, deltaNext } = await collectDelta(
        client.fetchJson,
        state.startUrl,
        { kind: "drive", id: state.resourceId },
        { includeDownloadUrl: state.includeDownloadUrl },
      );
      out.emit(buildDeltaBatch(schema, rows, deltaNext, state.includeDownloadUrl));
      state.done = true; // next process() call hits the done branch and finishes.
    },
  });
}

// ---------------------------------------------------------------------------
// sharepoint_list_items(site, list, delta_token?, expand?, page_size?)
//   (no include_download_url: list items are metadata, not file bytes — SPEC §4)
// ---------------------------------------------------------------------------

interface ListArgs {
  site: string;
  list: string;
  delta_token: string;
  expand: string;
  page_size: number;
}
interface ListState {
  done: boolean;
  startUrl: string;
  resourceId: string;
}

export function makeListItems(clientFactory: ClientFactory) {
  return defineTableFunction<ListArgs, ListState>({
    name: "sharepoint_list_items",
    description:
      "SharePoint listItems via Microsoft Graph per-list delta (incremental change feed with " +
      "removals; $expand=fields materializes user-defined columns). One delta_token per (site,list); " +
      "the {list -> deltaLink} map is held by the caller (multi-cursor).",
    args: {
      site: new Utf8(),
      list: new Utf8(),
      delta_token: new Utf8(),
      expand: new Utf8(),
      page_size: new Int64(),
    },
    argDefaults: {
      delta_token: "",
      expand: DEFAULT_EXPAND,
      page_size: DEFAULT_PAGE_SIZE,
    },
    argDocs: {
      site: "REQUIRED. The Microsoft Graph SharePoint site id whose list is scanned (the `siteId`).",
      list: "REQUIRED. The Microsoft Graph list id of the SharePoint list to scan (the `listId`), within `site`.",
      delta_token:
        "A previously persisted `@odata.deltaLink` (the `_delta_next` value from a prior scan's marker row for THIS list), replayed verbatim to return only what changed since. Empty (the default) performs a full sync.",
      expand:
        "Microsoft Graph `$expand` clause for the initial full sync. Defaults to `fields`, which materializes the list's user-defined columns (the item Title maps to the `name` column).",
      page_size:
        "Best-effort Microsoft Graph `$top` page size for the initial full sync (Graph may not honor it). Defaults to 200.",
    },
    examples: [
      {
        sql: "SELECT id, name, web_url, last_modified_date_time FROM azure.main.sharepoint_list_items('<siteId>', '<listId>') WHERE _row_kind IS NULL",
        description: "Full sync of a SharePoint list's items (data rows only)",
      },
      {
        sql: "SELECT id, change_type, name FROM azure.main.sharepoint_list_items('<siteId>', '<listId>', delta_token := '<@odata.deltaLink>')",
        description: "Incremental sync replaying a previously saved delta cursor for this list",
      },
      {
        sql: "SELECT _delta_next FROM azure.main.sharepoint_list_items('<siteId>', '<listId>') WHERE _row_kind = 'marker'",
        description: "Read the delta cursor to persist for this list's next sync",
      },
    ],
    tags: {
      "vgi.category": "list-items",
      "vgi.title": "SharePoint List Delta Feed",
      "vgi.keywords": JSON.stringify([
        "sharepoint",
        "list items",
        "listitem",
        "lists",
        "fields",
        "columns",
        "metadata",
        "delta",
        "change feed",
      ]),
      "vgi.doc_llm":
        "Incremental change feed of SharePoint list items (listItems) via Microsoft Graph per-list delta " +
        "query. Each scan returns items created or updated since the supplied delta_token (change_type " +
        "'upsert') and items deleted (change_type 'removed'), followed by a marker row whose _delta_next " +
        "is the cursor to persist for THIS list. Call with no delta_token for a full sync. `site` and " +
        "`list` are required (the Graph siteId and listId). `$expand=fields` materializes user-defined " +
        "columns; the item Title surfaces as the `name` column. List items are metadata, not file bytes, " +
        "so there is no download-url column.",
      "vgi.doc_md":
        "## sharepoint_list_items\n\n" +
        "SharePoint `listItems` as an incremental change feed backed by Microsoft Graph's " +
        "`sites/{site}/lists/{list}/items/delta` query. `site` and `list` are required. `$expand=fields` " +
        "(the default) materializes the list's user-defined columns — the item Title surfaces as `name`. " +
        "Read data rows with `WHERE _row_kind IS NULL`; take the next cursor from the marker row's " +
        "`_delta_next` and persist it per list. See the examples for runnable full-sync and incremental queries.",
      // Static: sharepoint_list_items never emits _download_url, so the schema is fixed (VGI307 static).
      "vgi.result_columns_schema": JSON.stringify([...BUSINESS_RESULT_COLUMNS, ...CONTROL_RESULT_COLUMNS]),
    },
    onBind: (p) => {
      if (!p.args.site) throw new Error("sharepoint_list_items: 'site' is required");
      if (!p.args.list) throw new Error("sharepoint_list_items: 'list' is required");
      return { outputSchema: fileItemSchema(false) }; // list items never emit _download_url
    },
    initialState: (p) => ({
      done: false,
      resourceId: `${p.args.site}/${p.args.list}`,
      startUrl: p.args.delta_token
        ? p.args.delta_token
        : listItemsStartUrl(p.args.site, p.args.list, p.args.expand, Number(p.args.page_size) || DEFAULT_PAGE_SIZE),
    }),
    process: async (p, state: ListState, out: OutputCollector) => {
      if (state.done) {
        out.finish();
        return;
      }
      const client = clientFactory(requireSecret(p, "sharepoint_list_items"));
      const schema = fileItemSchema(false);
      const { rows, deltaNext } = await collectDelta(
        client.fetchJson,
        state.startUrl,
        { kind: "list", id: state.resourceId },
        { includeDownloadUrl: false },
      );
      out.emit(buildDeltaBatch(schema, rows, deltaNext, false));
      state.done = true;
    },
  });
}

/** Build both table functions over one injected client factory. */
export function makeFunctions(clientFactory: ClientFactory) {
  return [makeDriveItems(clientFactory), makeListItems(clientFactory)];
}
