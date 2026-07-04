// The Graph content delta driver — pure logic over graph-core, no SDK / no network.
// ONE driver serves both driveItems (OneDrive / SharePoint document libraries) and
// listItems (SharePoint lists): a delta row is generic (id + change-type + a handful
// of business fields), so the same code walks either endpoint. This is the module the
// per-resource crash/resume + cursor-isolation proof (test/cursor-proof.test.ts)
// exercises — WITHOUT importing the VGI SDK, so it runs anywhere.
//
// Multi-cursor MAP note: each invocation manages EXACTLY ONE resource's cursor (one
// bare delta_token in / one deltaLink out). The map of {resourceId -> deltaLink} is
// the CALLER's to hold and fan out (SPEC §2b); the worker stays per-resource and
// stateless across the fleet, so every cursor is an independently-resumable atom.

import { paginate, isRemoved, ResyncRequired, type FetchJson } from "@vgi-azure/graph-core";

const GRAPH = "https://graph.microsoft.com/v1.0";

/** Graph's short-lived, pre-authed, NO-BEARER capability URL (SPEC §4a). A ~1h RBAC
 *  bypass if it leaks: gated behind include_download_url and NEVER logged. */
export const DOWNLOAD_URL_KEY = "@microsoft.graph.downloadUrl";

export type ResourceKind = "drive" | "list";
export type ChangeType = "upsert" | "removed";

/** Identity of the single resource a scan is a cursor over — carried into
 *  ResyncRequired so the caller can drop just THAT key from its map (SPEC §2b). */
export interface ResourceRef {
  kind: ResourceKind;
  id: string;
}

export interface FileRow {
  id: string;
  changeType: ChangeType;
  /** '@removed.reason' (changed|deleted) for removed rows, else null. */
  removedReason: string | null;
  name: string | null;
  size: number | null;
  webUrl: string | null;
  lastModified: string | null;
  /** Opt-in capability URL. ALWAYS null unless include_download_url is set AND this is
   *  a live file row (never folders / @removed tombstones) — SPEC §4a SECURITY. */
  downloadUrl: string | null;
}

export interface DeltaResult {
  rows: FileRow[];
  /** The verbatim @odata.deltaLink to persist as THIS resource's cursor next scan. */
  deltaNext: string;
}

export interface CollectOptions {
  /** Emit the download capability URL on file rows. Default false (SPEC §4a). */
  includeDownloadUrl?: boolean;
}

/** Initial full-sync URL for a drive. `$select`/`$top` are pinned for the life of the
 *  resulting delta token; on replay we follow the server's link VERBATIM (SPEC §2a). */
export function driveItemsStartUrl(drive: string, select: string, pageSize: number): string {
  return `${GRAPH}/drives/${encodeURIComponent(drive)}/root/delta` +
    `?$select=${encodeURIComponent(select)}&$top=${pageSize}`;
}

/** Initial full-sync URL for a SharePoint list. `$expand=fields` materializes the
 *  user-defined columns; `$select`/`$top` best-effort and pinned for the token life. */
export function listItemsStartUrl(site: string, list: string, expand: string, pageSize: number): string {
  return `${GRAPH}/sites/${encodeURIComponent(site)}/lists/${encodeURIComponent(list)}/items/delta` +
    `?$expand=${encodeURIComponent(expand)}&$top=${pageSize}`;
}

/** Map one raw Graph driveItem/listItem to a generic FileRow. `@removed` rows become
 *  id-only tombstones. downloadUrl is populated ONLY when opted in and present. */
export function toRow(obj: Record<string, unknown>, includeDownloadUrl: boolean): FileRow {
  const id = String(obj.id ?? "");
  if (isRemoved(obj)) {
    const removed = obj["@removed"] as { reason?: string } | undefined;
    return {
      id,
      changeType: "removed",
      removedReason: String(removed?.reason ?? "deleted"),
      name: null, size: null, webUrl: null, lastModified: null, downloadUrl: null,
    };
  }
  // listItems carry their business columns under `fields` (via $expand=fields);
  // driveItems carry name at top level. Fall back so one mapper serves both.
  const fields = obj.fields as Record<string, unknown> | undefined;
  const name =
    obj.name != null ? String(obj.name)
    : fields?.Title != null ? String(fields.Title)
    : null;
  const size = obj.size != null ? Number(obj.size) : null;
  const webUrl = obj.webUrl != null ? String(obj.webUrl) : null;
  const lastModified = obj.lastModifiedDateTime != null ? String(obj.lastModifiedDateTime) : null;
  const rawDownload = obj[DOWNLOAD_URL_KEY];
  const downloadUrl = includeDownloadUrl && rawDownload != null ? String(rawDownload) : null;
  return { id, changeType: "upsert", removedReason: null, name, size, webUrl, lastModified, downloadUrl };
}

/**
 * Drain ONE resource's delta scan to completion from `startUrl` (a fresh full-sync
 * URL, or a previously persisted @odata.deltaLink replayed VERBATIM). Returns every
 * change row plus the new deltaLink to persist for THAT resource.
 *
 * Loss-safety contract lives in the CALLER: persist `deltaNext` for this resource ONLY
 * after its rows are durably applied (SPEC §2a). On crash the caller still holds the
 * OLD token, re-runs this, and Graph replays the same window — absorbed idempotently by
 * upsert-by-id (@removed -> delete). Per-resource isolation: one poisoned cursor never
 * touches another's window.
 *
 * A 410 (resyncRequired) is re-thrown as a ResyncRequired carrying THIS resource's
 * identity, so the caller drops only this key from the map and full-re-syncs just it.
 */
export async function collectDelta(
  fetchJson: FetchJson,
  startUrl: string,
  resource: ResourceRef,
  opts: CollectOptions = {},
): Promise<DeltaResult> {
  const includeDownloadUrl = opts.includeDownloadUrl ?? false;
  const rows: FileRow[] = [];
  let deltaNext: string | undefined;
  try {
    for await (const page of paginate<Record<string, unknown>>(fetchJson, startUrl)) {
      for (const obj of page.value) rows.push(toRow(obj, includeDownloadUrl));
      if (page.deltaLink) deltaNext = page.deltaLink;
    }
  } catch (err) {
    if (err instanceof ResyncRequired) {
      // Re-key the resync onto THIS cursor so the caller drops one map entry, not all.
      throw new ResyncRequired(
        `resyncRequired kind=${resource.kind} id=${resource.id}: ${err.message}`,
      );
    }
    throw err;
  }
  if (deltaNext === undefined) {
    throw new Error(`delta scan for ${resource.kind}:${resource.id} ended without an @odata.deltaLink`);
  }
  return { rows, deltaNext };
}
