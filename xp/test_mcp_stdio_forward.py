from __future__ import print_function

import io
import json
import os
import socket
import tempfile
import threading
import unittest

from xp.mcp_diagnostics import DiagnosticLog
from xp.mcp_stdio_forward import BRIDGE_CONNECT_TIMEOUT_SECONDS
from xp.mcp_stdio_forward import connect_bridge
from xp.mcp_stdio_forward import forward


class FakeConnection(object):
    def __init__(self, responses=None, block_receive=False):
        self.responses = list(responses or [])
        self.block_receive = block_receive
        self.receive_released = threading.Event()
        self.sent = []
        self.shutdown_calls = []
        self.timeout_calls = []
        self.closed = False

    def recv(self, _size):
        if self.block_receive:
            self.receive_released.wait(1.0)
            return b""
        if self.responses:
            return self.responses.pop(0)
        return b""

    def sendall(self, data):
        self.sent.append(data)

    def settimeout(self, value):
        self.timeout_calls.append(value)

    def shutdown(self, how):
        self.shutdown_calls.append(how)
        if how == socket.SHUT_RDWR:
            self.receive_released.set()

    def close(self):
        self.closed = True
        self.receive_released.set()


class RecordingDiagnostics(object):
    def __init__(self):
        self.entries = []

    def emit(self, event, **fields):
        self.entries.append((event, fields))


class BridgeCleanupTests(unittest.TestCase):
    def test_connect_timeout_is_cleared_for_idle_session_receives(self):
        connection = FakeConnection()
        calls = []
        diagnostics = RecordingDiagnostics()

        def create_connection(address, timeout):
            calls.append((address, timeout))
            return connection

        connected = connect_bridge(
            "127.0.0.1",
            3211,
            diagnostics,
            create_connection,
        )

        self.assertIs(connection, connected)
        self.assertEqual(
            [(("127.0.0.1", 3211), BRIDGE_CONNECT_TIMEOUT_SECONDS)],
            calls,
        )
        self.assertEqual([None], connection.timeout_calls)
        self.assertIn(
            ("bridge_connect_complete", {"receive_mode": "blocking"}),
            diagnostics.entries,
        )

    def test_eof_half_closes_write_and_forwards_remaining_output(self):
        connection = FakeConnection([b'{"result":"ok"}\n', b""])
        output = io.BytesIO()

        forward(connection, io.BytesIO(b'{"method":"ping"}\n'), output, 0.05)

        self.assertEqual([b'{"method":"ping"}\n'], connection.sent)
        self.assertEqual(b'{"result":"ok"}\n', output.getvalue())
        self.assertEqual([socket.SHUT_WR], connection.shutdown_calls)
        self.assertTrue(connection.closed)

    def test_eof_forces_full_close_when_server_does_not_finish(self):
        connection = FakeConnection(block_receive=True)

        forward(connection, io.BytesIO(b""), io.BytesIO(), 0.01)

        self.assertEqual(
            [socket.SHUT_WR, socket.SHUT_RDWR],
            connection.shutdown_calls,
        )
        self.assertTrue(connection.closed)

    def test_bridge_diagnostics_exclude_tool_arguments(self):
        secret_text = "bridge payload must stay private"
        request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/call",
            "params": {
                "name": "clippy.speak",
                "arguments": {"text": secret_text},
            },
        }
        source = json.dumps(request).encode("utf-8") + b"\n"
        diagnostics = RecordingDiagnostics()
        connection = FakeConnection([b'{"jsonrpc":"2.0","id":1}\n', b""])

        forward(
            connection,
            io.BytesIO(source),
            io.BytesIO(),
            0.05,
            diagnostics,
        )

        self.assertIn("tools/call", repr(diagnostics.entries))
        self.assertIn("clippy.speak", repr(diagnostics.entries))
        self.assertNotIn(secret_text, repr(diagnostics.entries))
        self.assertIn(
            "bridge_stdout_write_complete",
            [event for event, _fields in diagnostics.entries],
        )

    def test_file_diagnostics_are_jsonl_and_best_effort(self):
        handle, path = tempfile.mkstemp()
        os.close(handle)
        try:
            diagnostics = DiagnosticLog(path)
            diagnostics.emit(
                "test_event",
                method="tools/list",
                detail="line one\nline two",
            )

            with open(path, "rb") as source:
                record = json.loads(source.read().decode("ascii"))

            self.assertEqual("test_event", record["event"])
            self.assertEqual("tools/list", record["method"])
            self.assertEqual("line one\\nline two", record["detail"])
            self.assertIn("timestamp", record)
            self.assertIn("pid", record)
        finally:
            os.unlink(path)

    def test_bridge_diagnostics_redact_unknown_routing_values(self):
        private_method = "method containing private caller text"
        source = json.dumps(
            {"jsonrpc": "2.0", "id": 1, "method": private_method}
        ).encode("utf-8") + b"\n"
        diagnostics = RecordingDiagnostics()
        connection = FakeConnection([])

        forward(
            connection,
            io.BytesIO(source),
            io.BytesIO(),
            0.05,
            diagnostics,
        )

        rendered = repr(diagnostics.entries)
        self.assertIn("'method': 'other'", rendered)
        self.assertNotIn(private_method, rendered)


if __name__ == "__main__":
    unittest.main()
