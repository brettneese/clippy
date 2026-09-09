from __future__ import print_function

import json
import os
import threading
import time


SERVER_LOG_PATH = r"C:\clippy\ClippyMcp.log"
BRIDGE_LOG_PATH = r"C:\clippy\ClippyMcpBridge.log"
MAX_FIELD_LENGTH = 160
SAFE_METHODS = frozenset(
    (
        "initialize",
        "notifications/initialized",
        "tools/list",
        "tools/call",
    )
)
SAFE_TOOLS = frozenset(
    (
        "clippy.animations",
        "clippy.show",
        "clippy.hide",
        "clippy.move",
        "clippy.speak",
        "clippy.think",
        "clippy.play",
    )
)


def _safe_field(value):
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if not isinstance(value, str):
        value = type(value).__name__
    return value.replace("\r", "\\r").replace("\n", "\\n")[:MAX_FIELD_LENGTH]


class DiagnosticLog(object):
    """Best-effort JSONL diagnostics that never write protocol payloads."""

    def __init__(self, path):
        self.path = path
        self._lock = threading.Lock()

    def emit(self, event, **fields):
        if not self.path:
            return

        record = {
            "timestamp": time.strftime(
                "%Y-%m-%dT%H:%M:%SZ",
                time.gmtime(),
            ),
            "pid": os.getpid(),
            "event": _safe_field(event),
        }
        for name, value in fields.items():
            record[_safe_field(name)] = _safe_field(value)

        data = json.dumps(
            record,
            ensure_ascii=True,
            separators=(",", ":"),
            sort_keys=True,
        ).encode("ascii") + b"\r\n"
        try:
            with self._lock:
                with open(self.path, "ab") as output:
                    output.write(data)
                    output.flush()
        except Exception:
            # Diagnostics must never break or contaminate MCP stdio.
            pass


def emit(diagnostics, event, **fields):
    if diagnostics is not None:
        diagnostics.emit(event, **fields)


def protocol_metadata(data):
    """Return safe routing metadata without retaining params or arguments."""
    metadata = {}
    try:
        if isinstance(data, bytes):
            data = data.decode("utf-8")
        request = json.loads(data)
    except (UnicodeDecodeError, ValueError, TypeError):
        metadata["message_shape"] = "invalid"
        return metadata

    if not isinstance(request, dict):
        metadata["message_shape"] = type(request).__name__
        return metadata

    metadata["message_shape"] = "object"
    method = request.get("method")
    if isinstance(method, str):
        metadata["method"] = method if method in SAFE_METHODS else "other"
    if method == "tools/call":
        params = request.get("params")
        if isinstance(params, dict) and isinstance(params.get("name"), str):
            tool = params["name"]
            metadata["tool"] = tool if tool in SAFE_TOOLS else "other"
    metadata["has_id"] = "id" in request
    return metadata
