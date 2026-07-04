// The VGI table functions: drive_items + list_items. Both are TABLE functions so
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
// list_items(site, list, delta_token?, expand?, page_size?)
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
    name: "list_items",
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
    onBind: (p) => {
      if (!p.args.site) throw new Error("list_items: 'site' is required");
      if (!p.args.list) throw new Error("list_items: 'list' is required");
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
      const client = clientFactory(requireSecret(p, "list_items"));
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
