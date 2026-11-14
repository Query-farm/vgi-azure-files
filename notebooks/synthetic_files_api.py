"""Synthetic paged Microsoft Graph drive snapshots and delta responses.

Envelope/cursor shapes follow the pinned worker's test/fake-files.ts (Query Farm,
MIT). File names, IDs, sizes and times are invented. No files are downloaded and
no request is forwarded to Microsoft Graph.
"""

from __future__ import annotations

import base64
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

BASE = "https://graph.microsoft.com/v1.0/drives/notebook-drive/root/delta"
MIB = 2**20


def item(identifier: str, name: str, mib: float | None, revised: bool = False) -> dict:
    """Define a synthetic file; an omitted size is unknown, not zero."""
    row = {
        "id": identifier,
        "name": name,
        "file": {},
        "webUrl": f"https://example.invalid/synthetic/{identifier}",
        "lastModifiedDateTime": "2026-06-12T12:00:00Z" if revised else "2026-06-10T12:00:00Z",
    }
    if mib is not None:
        row["size"] = int(mib * MIB)
    return row


INITIAL = [
    item("a", "report.pdf", 2),
    item("b", "simulation.bin", 24),
    item("c", "briefing.pptx", 8),
    item("d", "archive.zip", 32),
    item("e", "thumbnail.png", 0.5),
    item("f", "manifest.csv", None),
    item("g", "notes.txt", 0),
    item("h", "model.stl", 12),
]
CHANGES = [
    item("b", "simulation-final.bin", 40, True),
    item("c", "briefing-v2.pptx", 6, True),
    {"id": "d", "@removed": {"reason": "deleted"}},
    item("i", "recording.wav", 18, True),
]
CURRENT = [row for row in INITIAL if row["id"] not in {"b", "c", "d"}] + [
    row for row in CHANGES if "@removed" not in row
]


class SyntheticFilesApi:
    """Serve one explicit synthetic drive and record every requested URL."""

    def __init__(self) -> None:
        self.phase = 1
        self.requests: list[str] = []
        fixture = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, *args) -> None:
                pass

            def do_GET(self) -> None:  # noqa: N802
                fixture.requests.append(self.path)
                try:
                    payload, status = fixture.response(self.path), 200
                except (ValueError, KeyError) as error:
                    payload, status = {"error": str(error)}, 400
                data = json.dumps(payload).encode()
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

        self._server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)

    @property
    def url(self) -> str:
        return f"http://127.0.0.1:{self._server.server_port}"

    def start(self) -> SyntheticFilesApi:
        self._thread.start()
        return self

    def publish_changes(self) -> None:
        """Switch the local snapshot; this makes no request to a real service."""
        self.phase = 2

    def stop(self) -> None:
        self._server.shutdown()
        self._server.server_close()
        self._thread.join(timeout=2)

    def response(self, path: str) -> dict:
        parsed = urlparse(path)
        if parsed.path != "/v1.0/drives/notebook-drive/root/delta":
            raise ValueError("Only the synthetic notebook-drive delta endpoint is supported")
        query = parse_qs(parsed.query)
        offset = 0
        phase = self.phase
        if "$skiptoken" in query:
            raw = query["$skiptoken"][0]
            token = json.loads(base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4)))
            kind, phase, offset = token["kind"], token["phase"], token["offset"]
        elif "$deltatoken" in query:
            token = query["$deltatoken"][0]
            if token not in {"synthetic-v1", "synthetic-v2"} or query.get("opaque") != ["keep+exact"]:
                raise ValueError("Replay the synthetic deltaLink verbatim")
            kind = "changes" if token == "synthetic-v1" and phase == 2 else "empty"
        else:
            kind = "full"
        rows = (
            (INITIAL if phase == 1 else CURRENT) if kind == "full" else CHANGES if kind == "changes" else []
        )
        chunk = rows[offset : offset + 3]
        if offset + 3 < len(rows):
            token = base64.urlsafe_b64encode(
                json.dumps({"kind": kind, "phase": phase, "offset": offset + 3}).encode()
            )
            link = BASE + "?$skiptoken=" + token.decode().rstrip("=")
            return {"value": chunk, "@odata.nextLink": link}
        return {
            "value": chunk,
            "@odata.deltaLink": BASE + f"?$deltatoken=synthetic-v{phase}&opaque=keep%2Bexact",
        }
