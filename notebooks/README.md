# Microsoft 365 file inventory notebooks

[01 · Build the current file inventory](01_current_files_from_delta.ipynb) applies
a paged full scan and delta batch, including renames, resized files, an addition
and a removal. SQL compares the resulting inventory with a fresh full scan, then
measures the share of known bytes held by large files. The saved notebook contains
two figures. All file metadata and timestamps are synthetic.

## Install and run

Install [uv](https://docs.astral.sh/uv/), [Bun](https://bun.sh/) and Git. Validation
used Bun **1.3.14**. Copy this directory anywhere; no parent worker checkout or root
Python environment is needed. From this directory:

```sh
uv sync --locked
uv run --locked python execute.py --timeout 300
```

The separate lock pins Haybarn **2.0.0rc11**, JupySQL, plotting and execution tools.
The runner creates a temporary kernel with its interpreter and this directory as
cwd. It replaces the saved notebook atomically after success. Overrides require a
separate destination. For an existing Jupyter frontend:

```sh
uv run --locked python -m ipykernel install --user \
  --name vgi-azure-files-notebooks --display-name "Python (vgi-azure-files notebooks)"
```

Select that kernel, then **Restart Kernel → Run All** after editing parameters.
The runner does not register a global kernel. First runs require downloads.

## Two visible, portable connections

The native release attaches at the literal location
`github-auto://Query-farm/vgi-azure-files@v0.4.0` and supplies credential-free
`content_collections` discovery. Its analytical connection attaches at
`bash launch_fixture.sh`. This launcher clones the public remote revision
**4162efcdf50bd041b5baaa87faf41d895556b49c**, exactly the v0.4.0 source, into ignored
`artifacts/`, verifies the revision and unchanged source, and runs
`bun install --frozen-lockfile` there. Its lock pins `@query-farm/vgi 0.50.0`,
`@query-farm/vgi-rpc 0.29.7`, and Arrow 21.1.1. It does not discover or build the
parent checkout and does not change the user's root lock.

`fixture_worker.ts` imports the unchanged remote `makeFunctions` and `makeCatalog`
and registers them with the real VGI SDK. The existing client-factory seam supplies
a local HTTP adapter. The worker constructs URLs, follows pagination, decodes
business rows and tombstones, and emits its own Arrow schema and final marker.
The adapter maps only the expected Graph origin to loopback, rejects other origins,
and rejects writes. The SQL secret contains public dummy strings.

This tests the real exported functions with synthetic transport. It does not test
native-binary retrieval, Microsoft Entra authentication, tenant permissions,
Graph retry behavior, cursor expiration or a real Microsoft 365 account. No real
credentials are read, requests forwarded to Microsoft, files downloaded, content
inspected or remote state changed. The two connections are labelled separately.

## Focused experiment

```sh
uv run --locked python execute.py --timeout 300 \
  --parameters '{"LARGE_FILE_MIB": 16}' --output-dir artifacts/threshold-16
```

The default threshold is **8 MiB**, inclusive; accepted integers are **1–64**.
Predict fewer qualifying items and a smaller selected-byte total at **16 MiB**.
Compare the second figure's green bars and selected share. The inventory, first
figure and cursor checks should remain unchanged. Reset to **8** for the default.
Experiment copies carry this guide, runner, launcher, fixture, helper, environment
lock and branding assets. Generated source caches, logs and experiments are ignored.

## Fixture and inventory semantics

`synthetic_files_api.py` implements one loopback Graph drive-delta endpoint. The
response envelope and pagination follow `test/fake-files.ts` from the pinned MIT
source (Query Farm). It serves at most three rows per page and records request
paths. The notebook verifies that an opaque encoded cursor is replayed verbatim.
Unsupported endpoints and tokens fail explicitly; there is no live fallback.

The invented snapshots are dated **10 and 12 June 2026, 12:00 UTC**. Every object
is a file; folder aggregate sizes are outside this example. One file has unknown
size and one has a measured zero-byte size. MiB means 1,048,576 bytes. Byte totals
exclude unknowns and describe logical item sizes, not billed storage, transferred
bytes or reclaimed physical space.

A single final marker carries the per-drive deltaLink. It is control information,
not an item. The local key is `(drive_id, id)`, and the cursor is stored separately
for that drive. Removed rows contain only ID/reason; they delete local inventory
rows rather than inserting zero-size files. This simple merge explicitly requires
unique IDs in each change batch; ambiguous duplicates stop execution. A fresh
full scan supplies an independent same-snapshot reference after the merge.

All tables are in memory. The notebook demonstrates the application sequence, not
durable crash recovery. Production inventory changes and their cursor must commit
atomically on a single transactional connection. A real consumer must also handle
expired cursors and resynchronization. The large-file selection makes no content,
ownership, retention or deletion recommendation.

## Validation and upgrades

The source suite passes **11 Bun tests, 42 expectations**, and own-source
TypeScript checking. The CI-prescribed structural metadata check with vgi-lint
2.5.1 reports **100/100, L1 structural assurance, zero findings**; it intentionally
does not execute credentialed live examples. Notebook execution separately checks
seven SQL relationships, batch safety and opaque cursor replay through local HTTP.

Change both explicit release/revision pins when upgrading, inspect the remote
lock, then rerun default and experiment before replacing successful outputs. The
community VGI extension is downloaded independently of the notebook Python lock.
The unchanged local Strata Sun and VGI marks link to Query.Farm and VGI.

## SQL credential isolation

The notebook creates Haybarn with a fresh temporary `secret_directory` and
`allow_persistent_secrets=False` before attaching workers. It fully drains and
checks an empty secret inventory, then creates only explicitly `TEMPORARY`
public fixture secrets and verifies their `persistent=false`, `storage=memory`
metadata. Cleanup removes the temporary directory after closing the connection.
Existing user SQL secrets are neither read, modified nor deleted.
