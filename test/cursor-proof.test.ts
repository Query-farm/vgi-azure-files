// THE reason vgi-azure-files exists: prove the Graph content delta token is a durable,
// no-loss VGI scan cursor — AND that it holds PER RESOURCE across a multi-cursor map.
// This is the multi-cursor MAP proof (SPEC §2b, §7). It imports ONLY graph-core + this
// package's pure driver + bun:test (NO @query-farm/* SDK), so it runs anywhere.

import { test, expect } from "bun:test";
import { ResyncRequired } from "@vgi-azure/graph-core";
import { collectDelta, driveItemsStartUrl, listItemsStartUrl, type FileRow, type ResourceRef } from "../src/files-delta.js";
import { FakeGraphFiles, FakeResource } from "./fake-files.js";

const SELECT = "id,name,size,webUrl,lastModifiedDateTime";
const EXPAND = "fields";

const DRIVE_A: ResourceRef = { kind: "drive", id: "A" };
const LIST_X: ResourceRef = { kind: "list", id: "s/X" };

const driveStart = (pageSize = 100) => driveItemsStartUrl("A", SELECT, pageSize);
const listStart = (pageSize = 100) => listItemsStartUrl("s", "X", EXPAND, pageSize);

// --- a caller-side store, applied idempotently by id (the "downstream table") ---
type Store = Map<string, Partial<FileRow>>;
function apply(store: Store, rows: FileRow[]): void {
  for (const r of rows) {
    if (r.changeType === "removed") store.delete(r.id);
    else store.set(r.id, { name: r.name, size: r.size, webUrl: r.webUrl, lastModified: r.lastModified });
  }
}
const clone = (s: Store): Store => new Map([...s].map(([k, v]) => [k, { ...v }]));

function seedDrive(r: FakeResource): void {
  // A folder (no download URL) + two files (each carrying a capability URL).
  r.upsert({ id: "d-folder", name: "Docs", webUrl: "https://sp/Docs", lastModifiedDateTime: "2026-01-01T00:00:00Z", folder: { childCount: 2 } });
  r.upsert({ id: "d-file1", name: "a.txt", size: 11, webUrl: "https://sp/Docs/a.txt", lastModifiedDateTime: "2026-02-01T00:00:00Z", file: {}, "@microsoft.graph.downloadUrl": "https://dl.example/a?tok=1" });
  r.upsert({ id: "d-file2", name: "b.txt", size: 22, webUrl: "https://sp/Docs/b.txt", lastModifiedDateTime: "2026-02-02T00:00:00Z", file: {}, "@microsoft.graph.downloadUrl": "https://dl.example/b?tok=2" });
}
function seedList(r: FakeResource): void {
  r.upsert({ id: "x-item1", webUrl: "https://sp/lists/X/1", lastModifiedDateTime: "2026-03-01T00:00:00Z", fields: { Title: "Task 1" } });
  r.upsert({ id: "x-item2", webUrl: "https://sp/lists/X/2", lastModifiedDateTime: "2026-03-02T00:00:00Z", fields: { Title: "Task 2" } });
}

function newFleet(drivePageSize = 100) {
  const g = new FakeGraphFiles();
  const drive = g.add(new FakeResource("drives/A/root/delta", drivePageSize));
  const list = g.add(new FakeResource("sites/s/lists/X/items/delta"));
  seedDrive(drive);
  seedList(list);
  return { g, drive, list };
}

test("full sync pages through every driveItem/listItem and returns a verbatim deltaLink", async () => {
  const { g } = newFleet(/*pageSize*/ 2); // 3 drive objs → 2 pages
  let fetches = 0;
  const counting = async (url: string) => { fetches++; return g.fetch(url); };

  const a = await collectDelta(counting, driveStart(2), DRIVE_A);
  expect(a.rows.map((r) => r.id).sort()).toEqual(["d-file1", "d-file2", "d-folder"]);
  expect(a.rows.every((r) => r.changeType === "upsert")).toBe(true);
  expect(fetches).toBe(2); // two pages, nextLink followed verbatim
  // Verbatim replay: _delta_next is the FULL opaque @odata.deltaLink, not a bare token,
  // and points back at THIS drive's delta endpoint.
  expect(a.deltaNext).toContain("/drives/A/root/delta");
  expect(a.deltaNext).toContain("$deltatoken=");

  const x = await collectDelta(g.fetch, listStart(), LIST_X);
  expect(x.rows.map((r) => r.id).sort()).toEqual(["x-item1", "x-item2"]);
  expect(x.rows.find((r) => r.id === "x-item1")!.name).toBe("Task 1"); // $expand=fields → Title
  expect(x.deltaNext).toContain("/sites/s/lists/X/items/delta");
});

test("MULTI-CURSOR isolation: a drive mutation appears in the drive delta ONLY, never the list's", async () => {
  const { g, drive, list } = newFleet();
  const a0 = await collectDelta(g.fetch, driveStart(), DRIVE_A);
  const x0 = await collectDelta(g.fetch, listStart(), LIST_X);

  // Mutate ONLY drive A: add a file, remove a file.
  drive.upsert({ id: "d-file3", name: "c.txt", size: 33, webUrl: "https://sp/Docs/c.txt", lastModifiedDateTime: "2026-02-03T00:00:00Z", file: {}, "@microsoft.graph.downloadUrl": "https://dl.example/c?tok=3" });
  drive.remove("d-file1");

  const a1 = await collectDelta(g.fetch, a0.deltaNext, DRIVE_A);
  const aById = Object.fromEntries(a1.rows.map((r) => [r.id, r]));
  expect(Object.keys(aById).sort()).toEqual(["d-file1", "d-file3"]); // unchanged rows absent
  expect(aById["d-file3"]!.changeType).toBe("upsert");
  expect(aById["d-file1"]!.changeType).toBe("removed");
  expect(aById["d-file1"]!.removedReason).toBe("deleted");

  // X's cursor is untouched: replaying X0 yields NOTHING (drive's mutation is not here).
  const x1 = await collectDelta(g.fetch, x0.deltaNext, LIST_X);
  expect(x1.rows).toEqual([]);
  expect(x1.rows.some((r) => r.id.startsWith("d-"))).toBe(false);

  // Now mutate ONLY list X: X's delta shows exactly that; drive's ids never appear.
  list.upsert({ id: "x-item1", webUrl: "https://sp/lists/X/1", lastModifiedDateTime: "2026-03-09T00:00:00Z", fields: { Title: "Task 1 — CHANGED" } });
  const x2 = await collectDelta(g.fetch, x0.deltaNext, LIST_X);
  expect(x2.rows.map((r) => r.id)).toEqual(["x-item1"]);
  expect(x2.rows[0]!.name).toBe("Task 1 — CHANGED");
});

test("CRASH/RESUME per resource: replay after a crash yields identical state — no loss, no dup", async () => {
  const { g, drive } = newFleet(/*pageSize*/ 2);

  const a0 = await collectDelta(g.fetch, driveStart(2), DRIVE_A);
  const store: Store = new Map();
  apply(store, a0.rows);
  const A1 = a0.deltaNext; // durably persisted cursor for drive A

  // Source mutates.
  drive.upsert({ id: "d-file3", name: "c.txt", size: 33, webUrl: "https://sp/Docs/c.txt", lastModifiedDateTime: "2026-02-03T00:00:00Z", file: {}, "@microsoft.graph.downloadUrl": "https://dl.example/c?tok=3" });
  drive.remove("d-file1");
  drive.upsert({ id: "d-file2", name: "b2.txt", size: 222, webUrl: "https://sp/Docs/b.txt", lastModifiedDateTime: "2026-02-04T00:00:00Z", file: {}, "@microsoft.graph.downloadUrl": "https://dl.example/b?tok=9" });

  // Scan #2 runs, rows applied — but worker CRASHES before persisting the new cursor,
  // so the durable cursor for A is STILL A1.
  const a1 = await collectDelta(g.fetch, A1, DRIVE_A);
  const crashed = clone(store);
  apply(crashed, a1.rows);

  // Resume from the still-durable A1. Graph replays the SAME window for THIS drive.
  const a1b = await collectDelta(g.fetch, A1, DRIVE_A);
  const resumed = clone(store);
  apply(resumed, a1b.rows);

  expect(a1b.rows).toEqual(a1.rows); // identical replay
  expect([...resumed.entries()].sort()).toEqual([...crashed.entries()].sort());
  expect(resumed.has("d-file1")).toBe(false); // the delete survived
  expect(resumed.has("d-file3")).toBe(true);  // the add survived
  expect(resumed.get("d-file2")!.name).toBe("b2.txt"); // the change survived
});

test("PER-CURSOR resync: an expired drive token → ResyncRequired carrying the drive identity; the list never resyncs", async () => {
  const { g, drive, list } = newFleet();
  const a0 = await collectDelta(g.fetch, driveStart(), DRIVE_A);
  const x0 = await collectDelta(g.fetch, listStart(), LIST_X);

  drive.expireToken(a0.deltaNext); // only DRIVE A's token ages out

  const err = await collectDelta(g.fetch, a0.deltaNext, DRIVE_A).catch((e) => e);
  expect(err).toBeInstanceOf(ResyncRequired);
  expect((err as Error).message).toContain("kind=drive");
  expect((err as Error).message).toContain("id=A"); // identity so caller drops just this key

  // Recovery for A is a fresh full sync; X's cursor is untouched and still works.
  const recovered = await collectDelta(g.fetch, driveStart(), DRIVE_A);
  expect(recovered.rows.length).toBe(3);
  list.upsert({ id: "x-item2", webUrl: "https://sp/lists/X/2", lastModifiedDateTime: "2026-03-20T00:00:00Z", fields: { Title: "Task 2 — edited" } });
  const x1 = await collectDelta(g.fetch, x0.deltaNext, LIST_X);
  expect(x1.rows.map((r) => r.id)).toEqual(["x-item2"]); // list cursor never resynced
});

test("SECURITY: _download_url is ABSENT by default and PRESENT only when opted in (SPEC §4a)", async () => {
  const { g } = newFleet();

  // Default: include_download_url off → no row ever carries a capability URL.
  const off = await collectDelta(g.fetch, driveStart(), DRIVE_A, { includeDownloadUrl: false });
  expect(off.rows.every((r) => r.downloadUrl === null)).toBe(true);

  // Opted in: file rows carry the URL; the folder does NOT.
  const on = await collectDelta(g.fetch, driveStart(), DRIVE_A, { includeDownloadUrl: true });
  const byId = Object.fromEntries(on.rows.map((r) => [r.id, r]));
  expect(byId["d-file1"]!.downloadUrl).toBe("https://dl.example/a?tok=1");
  expect(byId["d-file2"]!.downloadUrl).toBe("https://dl.example/b?tok=2");
  expect(byId["d-folder"]!.downloadUrl).toBeNull(); // folders have no bytes to fetch
});

test("SECURITY: @removed tombstones never carry a download URL even when opted in", async () => {
  const { g, drive } = newFleet();
  const a0 = await collectDelta(g.fetch, driveStart(), DRIVE_A, { includeDownloadUrl: true });
  drive.remove("d-file1");
  const a1 = await collectDelta(g.fetch, a0.deltaNext, DRIVE_A, { includeDownloadUrl: true });
  const removed = a1.rows.find((r) => r.id === "d-file1")!;
  expect(removed.changeType).toBe("removed");
  expect(removed.downloadUrl).toBeNull();
});
