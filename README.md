# vgi-azure-files

A [VGI](https://query.farm) worker (**Microsoft 365 Content Delta**) for DuckDB. Microsoft 365 content as incremental (delta) DuckDB table functions over Microsoft Graph: `drive_items` (OneDrive / SharePoint document-library items, with an opt-in `_download_url` column) and `sharepoint_list_items` (SharePoint list items with their user-defined fields). Each scan returns the items created/updated or removed since the last cursor plus a single marker row (`_row_kind = 'marker'`) whose `_delta_next` holds the `@odata.deltaLink` to persist. Each call covers one drive or list; the caller keeps the `{resourceId -> deltaLink}` map. Call with no `delta_token` for a full sync. A credential-free `content_collections` view is also exposed.

## Quick start

Every [GitHub release](https://github.com/Query-farm/vgi-azure-files/releases) ships a prebuilt worker binary. The `vgi` extension downloads the one for your platform, checks its SHA-256 and caches it, so there is nothing to install besides DuckDB:

```sql
INSTALL vgi FROM community;
LOAD vgi;

ATTACH 'azure' AS azure (TYPE vgi, LOCATION 'github-auto://Query-farm/vgi-azure-files@v0.2.1');

-- Bind drive_items and expose its default result columns (credential-free; the 'drive'
-- placeholder satisfies onBind, LIMIT 0 skips the network. Supply a real driveId and
-- attach an azure_graph secret to sync).
SELECT id, name, size, web_url, last_modified_date_time FROM azure.main.drive_items('placeholder-drive-id') LIMIT 0;

-- Bind drive_items with include_download_url := true to expose the opt-in _download_url
-- capability column (credential-free; LIMIT 0 skips the network).
SELECT id, name, _download_url FROM azure.main.drive_items('placeholder-drive-id', include_download_url := true) LIMIT 0;
```

The examples are credential-free `LIMIT 0` bind probes: they check that each function binds and show its result columns without calling Azure. To pull real rows, create an `azure_graph` secret (see [Credentials](#credentials)) and drop the `LIMIT 0`.

All the Query Farm Azure workers use the catalog name `azure`. To attach more than one in the same session, give each its own alias (for example `ATTACH 'azure' AS azure_files (...)`) and qualify queries with that alias instead of `azure`.

Prebuilt binaries cover macOS (Apple silicon), Linux (x86-64 and arm64). `github-auto://` needs an explicit release tag; change `@v0.2.1` to move to a newer release.

### Other ways to run it

- **Docker** (the image bundles the worker; needs a container runtime):

  ```sql
  ATTACH 'azure' AS azure (TYPE vgi, LOCATION 'docker://ghcr.io/query-farm/vgi-azure-files:0.2.1');
  ```

- **Manual download.** Download the archive for your platform from the [releases page](https://github.com/Query-farm/vgi-azure-files/releases), verify it against its `.sha256` file (a cosign `.cosign.bundle` is published alongside), unpack it, and set `LOCATION` to the path of the extracted worker executable.

### Build from source

To build and test the worker yourself, see [Development](#development).

## Credentials

Live queries authenticate with an `azure_graph` secret holding app-only Microsoft Entra client credentials for a service principal: `TENANT_ID`, `CLIENT_ID` and `CLIENT_SECRET` (the client secret is redacted when secrets are listed). The worker registers the `azure_graph` secret type when it is attached, so create the secret after the `ATTACH`:

```sql
ATTACH 'azure' AS azure (TYPE vgi, LOCATION 'github-auto://Query-farm/vgi-azure-files@v0.2.1');
CREATE SECRET azure_sp (TYPE azure_graph, TENANT_ID '<tenant-guid>', CLIENT_ID '<app-guid>', CLIENT_SECRET '<secret>');
```

For this worker the service principal needs Microsoft Graph application read permission for the resources being queried (for example `Sites.Selected`, `Sites.Read.All` or `Files.Read.All`).

## Development

Requires [Bun](https://bun.sh). `bun install` also fetches the shared `@vgi-azure/graph-core` package from GitHub.

```bash
bun install && bun test
bun run typecheck
```

To attach a local checkout, point `LOCATION` at the launch wrapper:

```sql
ATTACH 'azure' AS azure (TYPE vgi, LOCATION '/path/to/vgi-azure-files/bin/vgi-azure-files-worker');
```

## License

MIT. See [LICENSE](LICENSE).
