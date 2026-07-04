import { test, expect } from "bun:test";
import { fileItemSchema, buildDeltaBatch, DOWNLOAD_URL_COL } from "../src/schema.js";
import type { FileRow } from "../src/files-delta.js";

const DELTA = "https://graph.microsoft.com/v1.0/drives/A/root/delta?$deltatoken=Z";

const fileRow = (id: string, dl: string | null): FileRow => ({
  id, changeType: "upsert", removedReason: null,
  name: `${id}.txt`, size: 10, webUrl: `https://sp/${id}`, lastModified: "2026-02-01T00:00:00Z",
  downloadUrl: dl,
});

test("default schema OMITS _download_url; opting in APPENDS it (SECURITY §4a)", () => {
  expect(fileItemSchema(false).fields.map((f) => f.name)).toEqual([
    "id", "change_type", "removed_reason", "name", "size", "web_url", "last_modified_date_time", "_row_kind", "_delta_next",
  ]);
  const cols = fileItemSchema(true).fields.map((f) => f.name);
  expect(cols).toContain(DOWNLOAD_URL_COL);
  // control cols stay last, capability URL sits just before them
  expect(cols.slice(-3)).toEqual([DOWNLOAD_URL_COL, "_row_kind", "_delta_next"]);
});

test("buildDeltaBatch: N data rows + exactly 1 marker row carrying the verbatim deltaLink", () => {
  const schema = fileItemSchema(false);
  const rows = [fileRow("d1", null), { id: "d2", changeType: "removed", removedReason: "deleted", name: null, size: null, webUrl: null, lastModified: null, downloadUrl: null } as FileRow];
  const batch = buildDeltaBatch(schema, rows, DELTA, false) as { numRows: number };
  expect(batch.numRows).toBe(3); // 2 data + 1 marker
});

test("buildDeltaBatch on an empty delta still emits the marker row", () => {
  const schema = fileItemSchema(false);
  const batch = buildDeltaBatch(schema, [], DELTA, false) as { numRows: number };
  expect(batch.numRows).toBe(1);
});

test("buildDeltaBatch with opt-in emits the _download_url column shape", () => {
  const schema = fileItemSchema(true);
  const batch = buildDeltaBatch(schema, [fileRow("d1", "https://dl.example/x")], DELTA, true) as { numRows: number };
  expect(batch.numRows).toBe(2); // 1 data + 1 marker
});
