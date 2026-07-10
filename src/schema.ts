// Arrow output schema + row->batch mapping for the content delta functions.
// Both drive_items and sharepoint_list_items share ONE column set (id + change cols + a handful
// of business fields), so the schema is data, not duplicated code.
//
// SECURITY (SPEC §4a): the `_download_url` column is GATED. It is omitted from the
// schema entirely unless include_download_url := true — a short-lived, pre-authed,
// no-bearer capability URL must never be a default column.

import { Schema, Field, Utf8, Int64 } from "@query-farm/apache-arrow";
import { batchFromColumns } from "@query-farm/vgi";
import { ROW_KIND, MARKER, DELTA_NEXT } from "@vgi-azure/graph-core";
import type { FileRow } from "./files-delta.js";

/** Opt-in-only capability-URL column (SPEC §4a). */
export const DOWNLOAD_URL_COL = "_download_url";

/** Business columns (never null-carrying control cols). Shared by both functions. */
const BUSINESS_COLS = ["id", "change_type", "removed_reason", "name", "size", "web_url", "last_modified_date_time"] as const;

/**
 * Build the output schema. `_download_url` is appended ONLY when opted in; the default
 * schema omits it (SPEC §4a). Control columns `_row_kind` / `_delta_next` come last per
 * the graph-core marker-row contract.
 */
export function fileItemSchema(includeDownloadUrl: boolean): Schema {
  const fields = [
    new Field("id", new Utf8(), true),
    new Field("change_type", new Utf8(), true),
    new Field("removed_reason", new Utf8(), true),
    new Field("name", new Utf8(), true),
    new Field("size", new Int64(), true),
    new Field("web_url", new Utf8(), true),
    new Field("last_modified_date_time", new Utf8(), true),
  ];
  if (includeDownloadUrl) fields.push(new Field(DOWNLOAD_URL_COL, new Utf8(), true));
  fields.push(new Field(ROW_KIND, new Utf8(), true));
  fields.push(new Field(DELTA_NEXT, new Utf8(), true));
  return new Schema(fields);
}

/**
 * Build one Arrow batch: N change rows (`_row_kind` null) followed by EXACTLY ONE
 * marker row (all business columns null, `_row_kind='marker'`, `_delta_next` = the
 * verbatim deltaLink for this resource). Consumers read data via `WHERE _row_kind IS
 * NULL` and the cursor from the single marker row — the graph-core §D contract. N+1
 * rows in one batch keeps the marker atomic with its data.
 */
export function buildDeltaBatch(
  schema: Schema,
  rows: FileRow[],
  deltaNext: string,
  includeDownloadUrl: boolean,
) {
  const cols: Record<string, unknown[]> = {};
  for (const c of BUSINESS_COLS) cols[c] = [];
  if (includeDownloadUrl) cols[DOWNLOAD_URL_COL] = [];
  cols[ROW_KIND] = [];
  cols[DELTA_NEXT] = [];

  for (const r of rows) {
    cols.id!.push(r.id);
    cols.change_type!.push(r.changeType);
    cols.removed_reason!.push(r.removedReason);
    cols.name!.push(r.name);
    // Arrow Int64 rich value is a plain bigint (or null). Sizes are integral bytes.
    cols.size!.push(r.size == null ? null : BigInt(Math.trunc(r.size)));
    cols.web_url!.push(r.webUrl);
    cols.last_modified_date_time!.push(r.lastModified);
    if (includeDownloadUrl) cols[DOWNLOAD_URL_COL]!.push(r.downloadUrl);
    cols[ROW_KIND]!.push(null);
    cols[DELTA_NEXT]!.push(null);
  }

  // The single strict marker row: all business columns null, cursor carried on it.
  cols.id!.push(null);
  cols.change_type!.push(null);
  cols.removed_reason!.push(null);
  cols.name!.push(null);
  cols.size!.push(null);
  cols.web_url!.push(null);
  cols.last_modified_date_time!.push(null);
  if (includeDownloadUrl) cols[DOWNLOAD_URL_COL]!.push(null);
  cols[ROW_KIND]!.push(MARKER);
  cols[DELTA_NEXT]!.push(deltaNext);

  return batchFromColumns(cols as Record<string, unknown[]>, schema);
}
