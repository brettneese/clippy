from __future__ import print_function

import io
import json
import socket
import unittest

from xp.clippy_agent import (
    ClippyController,
    JsonRpcServer,
    InvalidParams,
    McpServer,
    MCP_PROTOCOL_VERSION,
    MCP_SUPPORTED_PROTOCOL_VERSIONS,
    serve,
    serve_mcp,
    serve_mcp_connection,
)


class FakeBalloon(object):
    def __init__(self):
        self.Style = None


class FakeCharacter(object):
    def __init__(self):
        self.AnimationNames = ["Greeting", "Thinking", "Wave"]
        self.Balloon = FakeBalloon()
        self.calls = []

    def Show(self):
        self.calls.append(("show",))

    def Hide(self):
        self.calls.append(("hide",))

    def MoveTo(self, x, y):
        self.calls.append(("move", x, y))

    def Speak(self, text):
        self.calls.append(("speak", text))

    def Think(self, text):
        self.calls.append(("think", text))

    def Play(self, animation):
        self.calls.append(("play", animation))

    def StopAll(self):
        self.calls.append(("stop_all",))


class FakeCharacters(object):
    def __init__(self, character):
        self.character = character
        self.loaded = []
        self.unloaded = []

    def Load(self, key, path):
        self.loaded.append((key, path))

    def Character(self, key):
        return self.character

    def Unload(self, key):
        self.unloaded.append(key)


class FakeAgent(object):
    def __init__(self):
        self.Connected = False
        self.character = FakeCharacter()
        self.Characters = FakeCharacters(self.character)


class PumpingController(ClippyController):
    def __init__(self, agent_factory):
        ClippyController.__init__(self, agent_factory)
        self.pump_count = 0

    def pump_messages(self):
        self.pump_count += 1


class FakeSocketConnection(object):
    def __init__(self, source):
        self.chunks = [source, b""]
        self.output = []
        self.shutdown_calls = []
        self.closed = False

    def recv(self, _size):
        return self.chunks.pop(0)

    def sendall(self, data):
        self.output.append(data)

    def shutdown(self, how):
        self.shutdown_calls.append(how)

    def close(self):
        self.closed = True


class RecordingDiagnostics(object):
    def __init__(self):
        self.entries = []

    def emit(self, event, **fields):
        self.entries.append((event, fields))


class ClippyControllerTests(unittest.TestCase):
    def setUp(self):
        self.agent = FakeAgent()
        self.controller = ClippyController(lambda: self.agent)
        self.controller.connect()

    def tearDown(self):
        self.controller.close()

    def test_connects_once_and_enumerates_runtime_animations(self):
        self.controller.connect()
        self.assertTrue(self.agent.Connected)
        self.assertEqual(7, self.agent.character.Balloon.Style)
        self.assertEqual(
            ["Greeting", "Thinking", "Wave"],
            self.controller.animation_names(),
        )
        self.assertEqual(1, len(self.agent.Characters.loaded))

    def test_exposes_only_clippy_actions(self):
        server = JsonRpcServer(self.controller)
        server.handle({"jsonrpc": "2.0", "method": "clippy.show"})
        server.handle(
            {
                "jsonrpc": "2.0",
                "method": "clippy.move",
                "params": {"x": 320, "y": 240},
            }
        )
        server.handle(
            {
                "jsonrpc": "2.0",
                "method": "clippy.speak",
                "params": {"text": "Hello from XP"},
            }
        )
        server.handle(
            {
                "jsonrpc": "2.0",
                "method": "clippy.think",
                "params": {"text": "Still thinking"},
            }
        )
        server.handle(
            {
                "jsonrpc": "2.0",
                "method": "clippy.play",
                "params": {"animation": "Wave"},
            }
        )
        server.handle({"jsonrpc": "2.0", "method": "clippy.hide"})

        self.assertEqual(
            [
                ("show",),
                ("move", 320, 240),
                ("speak", "Hello from XP"),
                ("think", "Still thinking"),
                ("play", "Wave"),
                ("hide",),
            ],
            self.agent.character.calls,
        )

    def test_rejects_animation_not_reported_by_character(self):
        with self.assertRaises(InvalidParams):
            self.controller.play("DefinitelyNotInstalled")
        self.assertNotIn(
            ("play", "DefinitelyNotInstalled"), self.agent.character.calls
        )

    def test_rejects_extra_params_and_invalid_coordinates(self):
        server = JsonRpcServer(self.controller)
        with self.assertRaises(InvalidParams):
            server.handle(
                {
                    "jsonrpc": "2.0",
                    "method": "clippy.show",
                    "params": {"command": "cmd.exe"},
                }
            )
        with self.assertRaises(InvalidParams):
            self.controller.move(True, 20)


class JsonRpcServerTests(unittest.TestCase):
    def test_stdio_session_keeps_one_controller_and_returns_json_rpc(self):
        agent = FakeAgent()
        controller = PumpingController(lambda: agent)
        requests = [
            {"jsonrpc": "2.0", "id": 1, "method": "clippy.animations"},
            {"jsonrpc": "2.0", "id": 2, "method": "clippy.show"},
            {
                "jsonrpc": "2.0",
                "id": 3,
                "method": "clippy.play",
                "params": {"animation": "Greeting"},
            },
            {"jsonrpc": "2.0", "id": 4, "method": "clippy.shutdown"},
        ]
        source = b"".join(
            json.dumps(request).encode("utf-8") + b"\n" for request in requests
        )
        output = io.BytesIO()

        serve(io.BytesIO(source), output, controller)

        responses = [
            json.loads(line.decode("utf-8"))
            for line in output.getvalue().splitlines()
        ]
        self.assertEqual([1, 2, 3, 4], [item["id"] for item in responses])
        self.assertEqual(
            ["Greeting", "Thinking", "Wave"],
            responses[0]["result"]["animations"],
        )
        self.assertEqual(1, len(agent.Characters.loaded))
        self.assertEqual(["Clippy"], agent.Characters.unloaded)
        self.assertGreater(controller.pump_count, 0)

    def test_invalid_animation_returns_invalid_params(self):
        agent = FakeAgent()
        controller = ClippyController(lambda: agent)
        request = {
            "jsonrpc": "2.0",
            "id": "bad-animation",
            "method": "clippy.play",
            "params": {"animation": "Dance"},
        }
        output = io.BytesIO()

        serve(
            io.BytesIO(json.dumps(request).encode("utf-8") + b"\n"),
            output,
            controller,
        )

        response = json.loads(output.getvalue().decode("utf-8"))
        self.assertEqual(-32602, response["error"]["code"])
        self.assertIn("CLIPPIT.ACS", response["error"]["message"])

    def test_invalid_request_and_explicit_null_id_receive_errors(self):
        agent = FakeAgent()
        controller = ClippyController(lambda: agent)
        requests = [
            [],
            {"jsonrpc": "2.0", "id": None, "method": 42},
        ]
        source = b"".join(
            json.dumps(request).encode("utf-8") + b"\n" for request in requests
        )
        output = io.BytesIO()

        serve(io.BytesIO(source), output, controller)

        responses = [
            json.loads(line.decode("utf-8"))
            for line in output.getvalue().splitlines()
        ]
        self.assertEqual(-32600, responses[0]["error"]["code"])
        self.assertEqual(-32600, responses[1]["error"]["code"])
        self.assertIsNone(responses[0]["id"])
        self.assertIsNone(responses[1]["id"])


class McpServerTests(unittest.TestCase):
    def _initialize(self):
        return {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": MCP_PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {"name": "fixture", "version": "1.0"},
            },
        }

    def test_lifecycle_tools_and_notifications_use_mcp_shapes(self):
        agent = FakeAgent()
        controller = PumpingController(lambda: agent)
        requests = [
            self._initialize(),
            {"jsonrpc": "2.0", "method": "notifications/initialized"},
            {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "tools/list",
                "params": {"_meta": {"progressToken": 0}},
            },
            {
                "jsonrpc": "2.0",
                "id": 3,
                "method": "tools/call",
                "params": {
                    "name": "clippy.speak",
                    "arguments": {"text": "Hello"},
                    "_meta": {"progressToken": 1},
                },
            },
            {
                "jsonrpc": "2.0",
                "id": 4,
                "method": "tools/call",
                "params": {"name": "clippy.animations", "arguments": {}},
            },
        ]
        source = b"".join(
            json.dumps(request).encode("utf-8") + b"\n" for request in requests
        )
        output = io.BytesIO()

        serve_mcp(io.BytesIO(source), output, controller)

        responses = [
            json.loads(line.decode("utf-8"))
            for line in output.getvalue().splitlines()
        ]
        self.assertEqual([1, 2, 3, 4], [item["id"] for item in responses])
        self.assertEqual(MCP_PROTOCOL_VERSION, responses[0]["result"]["protocolVersion"])
        self.assertEqual({"tools": {}}, responses[0]["result"]["capabilities"])
        self.assertEqual(
            [
                "clippy.animations",
                "clippy.show",
                "clippy.hide",
                "clippy.move",
                "clippy.speak",
                "clippy.think",
                "clippy.play",
            ],
            [tool["name"] for tool in responses[1]["result"]["tools"]],
        )
        self.assertEqual(
            {"queued": True}, responses[2]["result"]["structuredContent"]
        )
        self.assertEqual(
            ["Greeting", "Thinking", "Wave"],
            responses[3]["result"]["structuredContent"]["animations"],
        )
        self.assertEqual([("speak", "Hello")], agent.character.calls[:1])
        self.assertIn(("hide",), agent.character.calls)
        self.assertEqual(["Clippy"], agent.Characters.unloaded)

    def test_tool_failures_are_tool_results_and_never_reach_com(self):
        agent = FakeAgent()
        controller = ClippyController(lambda: agent)
        server = McpServer(controller)
        server.handle(self._initialize())
        server.handle({"jsonrpc": "2.0", "method": "notifications/initialized"})

        response = server.handle(
            {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "tools/call",
                "params": {
                    "name": "clippy.play",
                    "arguments": {"animation": "NotInstalled"},
                },
            }
        )

        self.assertTrue(response["isError"])
        self.assertEqual(
            "The requested animation is not installed in Clippit.",
            response["structuredContent"]["error"],
        )
        self.assertEqual([], agent.character.calls)
        controller.close()

    def test_unsupported_version_is_safe_and_does_not_connect(self):
        agent = FakeAgent()
        controller = ClippyController(lambda: agent)
        request = self._initialize()
        request["params"]["protocolVersion"] = "1999-01-01"
        output = io.BytesIO()

        serve_mcp(
            io.BytesIO(json.dumps(request).encode("utf-8") + b"\n"),
            output,
            controller,
        )

        response = json.loads(output.getvalue().decode("utf-8"))
        self.assertEqual(-32602, response["error"]["code"])
        for version in MCP_SUPPORTED_PROTOCOL_VERSIONS:
            self.assertIn(version, response["error"]["message"])
        self.assertNotIn("CLIPPIT.ACS", response["error"]["message"])
        self.assertEqual([], agent.Characters.loaded)

    def test_codex_protocol_version_is_negotiated(self):
        agent = FakeAgent()
        controller = PumpingController(lambda: agent)
        request = self._initialize()
        request["params"]["protocolVersion"] = "2025-06-18"
        output = io.BytesIO()

        serve_mcp(
            io.BytesIO(json.dumps(request).encode("utf-8") + b"\n"),
            output,
            controller,
        )

        response = json.loads(output.getvalue().decode("utf-8"))
        self.assertEqual(
            "2025-06-18", response["result"]["protocolVersion"]
        )
        self.assertEqual(
            [("Clippy", r"C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS")],
            agent.Characters.loaded,
        )

    def test_malformed_and_preinitialized_requests_receive_protocol_errors(self):
        agent = FakeAgent()
        controller = ClippyController(lambda: agent)
        requests = [
            {"jsonrpc": "2.0", "id": 1, "method": "tools/list"},
            [],
        ]
        source = b"".join(
            json.dumps(request).encode("utf-8") + b"\n" for request in requests
        )
        output = io.BytesIO()

        serve_mcp(io.BytesIO(source), output, controller)

        responses = [
            json.loads(line.decode("utf-8"))
            for line in output.getvalue().splitlines()
        ]
        self.assertEqual(-32002, responses[0]["error"]["code"])
        self.assertEqual(-32600, responses[1]["error"]["code"])

    def test_socket_disconnect_closes_client_and_preserves_controller(self):
        agent = FakeAgent()
        controller = ClippyController(lambda: agent)
        requests = [
            self._initialize(),
            {"jsonrpc": "2.0", "method": "notifications/initialized"},
            {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
        ]
        source = b"".join(
            json.dumps(request).encode("utf-8") + b"\n" for request in requests
        )
        first = FakeSocketConnection(source)
        second = FakeSocketConnection(source)

        serve_mcp_connection(first, controller)
        serve_mcp_connection(second, controller)

        self.assertTrue(first.closed)
        self.assertTrue(second.closed)
        self.assertEqual([socket.SHUT_RDWR], first.shutdown_calls)
        self.assertEqual([socket.SHUT_RDWR], second.shutdown_calls)
        self.assertEqual(1, len(agent.Characters.loaded))
        self.assertEqual([], agent.Characters.unloaded)

        controller.close()
        self.assertEqual(["Clippy"], agent.Characters.unloaded)

    def test_diagnostics_record_routing_without_tool_arguments(self):
        agent = FakeAgent()
        controller = PumpingController(lambda: agent)
        diagnostics = RecordingDiagnostics()
        secret_text = "diagnostic payload must stay private"
        requests = [
            self._initialize(),
            {"jsonrpc": "2.0", "method": "notifications/initialized"},
            {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "tools/call",
                "params": {
                    "name": "clippy.speak",
                    "arguments": {"text": secret_text},
                },
            },
        ]
        source = b"".join(
            json.dumps(request).encode("utf-8") + b"\n" for request in requests
        )

        serve_mcp(
            io.BytesIO(source),
            io.BytesIO(),
            controller,
            diagnostics=diagnostics,
            session_id=7,
        )

        events = [event for event, _fields in diagnostics.entries]
        self.assertIn("controller_connect_complete", events)
        self.assertIn("tool_call_start", events)
        self.assertIn("tool_call_complete", events)
        self.assertIn("mcp_response_write_complete", events)
        self.assertIn("mcp_session_close_complete", events)
        self.assertIn("clippy.speak", repr(diagnostics.entries))
        self.assertNotIn(secret_text, repr(diagnostics.entries))


if __name__ == "__main__":
    unittest.main()
