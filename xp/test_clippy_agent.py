from __future__ import print_function

import io
import json
import unittest

from xp.clippy_agent import ClippyController, JsonRpcServer, InvalidParams, serve


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


if __name__ == "__main__":
    unittest.main()
