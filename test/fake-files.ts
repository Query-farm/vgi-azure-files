// A stateful fake of Microsoft Graph's content /delta endpoints — enough to prove the
// MULTI-CURSOR contract: N independent resources (drives + lists), each with its own
// versioned snapshots, real add/change/delete diffs, multi-page paging via $skiptoken,
// verbatim deltaLink, and per-resource 410 resyncRequired. No network.
//
// Each FakeResource is a directory-style delta store bound to one Graph base PATH; the
// FakeGraphFiles router dispatches a URL to the matching resource by pathname, so drive
// A's mutations can never leak into list X's delta window — the isolation the proof needs.

import { ResyncRequired, type FetchJson } from "@vgi-azure/graph-core";

const GRAPH = "https://graph.microsoft.com/v1.0";

export interface FakeObj {
  id: string;
  [k: string]: unknown;
}

function b64(s: string): string {
  return Buffer.from(s, "utf8").toString("base64url");
}
function unb64(s: string): string {
  return Buffer.from(s, "base64url").toString("utf8");
}

/** One delta cursor's worth of state, bound to a Graph base path like
 *  `drives/A/root/delta` or `sites/s/lists/X/items/delta`. */
export class FakeResource {
  private objs = new Map<string, FakeObj>();
  private version = 0;
  private snapshots = new Map<number, Map<string, FakeObj>>();
  private expired = new Set<string>();

  constructor(
    /** Path suffix (no leading slash, no query) this resource answers for. */
    readonly basePath: string,
    private readonly pageSize: number = 100,
  ) {
    this.snap();
  }

  private snap(): void {
    this.snapshots.set(this.version, new Map([...this.objs].map(([k, v]) => [k, { ...v }])));
  }
  private bump(): void {
    this.version++;
    this.snap();
  }

  upsert(o: FakeObj): void {
    this.objs.set(o.id, { ...o });
    this.bump();
  }
  remove(id: string): void {
    if (this.objs.delete(id)) this.bump();
  }
  /** Mark a delta token (version string) expired → next replay 410s (this resource only). */
  expireToken(deltaLinkUrl: string): void {
    const tok = new URL(deltaLinkUrl).searchParams.get("$deltatoken");
    if (tok) this.expired.add(tok);
  }

  private changeList(fromVersion: number | null): Record<string, unknown>[] {
    if (fromVersion === null) {
      return [...this.objs.values()].map((o) => ({ ...o }));
    }
    const base = this.snapshots.get(fromVersion) ?? new Map<string, FakeObj>();
    const out: Record<string, unknown>[] = [];
    for (const [id, cur] of this.objs) {
      const prev = base.get(id);
      if (!prev || JSON.stringify(prev) !== JSON.stringify(cur)) out.push({ ...cur });
    }
    for (const id of base.keys()) {
      if (!this.objs.has(id)) out.push({ id, "@removed": { reason: "deleted" } });
    }
    return out;
  }

  handle(u: URL): Record<string, unknown> {
    const dt = u.searchParams.get("$deltatoken");
    const sk = u.searchParams.get("$skiptoken");

    if (dt && this.expired.has(dt)) {
      throw new ResyncRequired(`resyncRequired: token ${dt} expired on ${this.basePath}`);
    }

    let base: number | null;
    let offset: number;
    if (sk) {
      const parsed = JSON.parse(unb64(sk)) as { base: number | null; offset: number };
      base = parsed.base;
      offset = parsed.offset;
    } else if (dt) {
      base = Number(dt);
      offset = 0;
    } else {
      base = null; // full sync
      offset = 0;
    }

    const all = this.changeList(base);
    const slice = all.slice(offset, offset + this.pageSize);
    const nextOffset = offset + this.pageSize;

    if (nextOffset < all.length) {
      const skiptoken = b64(JSON.stringify({ base, offset: nextOffset }));
      return { value: slice, "@odata.nextLink": `${GRAPH}/${this.basePath}?$skiptoken=${skiptoken}` };
    }
    // Final page → deltaLink carrying the current version as the token (verbatim replay).
    return { value: slice, "@odata.deltaLink": `${GRAPH}/${this.basePath}?$deltatoken=${this.version}` };
  }
}

/** Routes a Graph URL to the FakeResource whose basePath its pathname ends with. */
export class FakeGraphFiles {
  private readonly resources: FakeResource[] = [];

  add(r: FakeResource): FakeResource {
    this.resources.push(r);
    return r;
  }

  private route(pathname: string): FakeResource {
    // Longest basePath first so a more specific path wins.
    const match = [...this.resources]
      .sort((a, b) => b.basePath.length - a.basePath.length)
      .find((r) => pathname.endsWith("/" + r.basePath));
    if (!match) throw new Error(`fake: no resource for ${pathname}`);
    return match;
  }

  fetch: FetchJson = async (url) => {
    const u = new URL(url);
    return this.route(u.pathname).handle(u);
  };
}
