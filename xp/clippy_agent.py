from __future__ import print_function

import json
import sys


CHARACTER_KEY = "Clippy"
CHARACTER_PATH = (
    r"C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS"
)
MAX_TEXT_LENGTH = 2000
MIN_COORDINATE = -32768
MAX_COORDINATE = 32767


class ControllerError(Exception):
    pass


class InvalidParams(ControllerError):
    pass


class InvalidRequest(ControllerError):
    pass


class MethodNotFound(ControllerError):
    pass


def create_agent_control():
    try:
        import win32com.client
    except ImportError:
        raise ControllerError(
            "pywin32 build 220 for 32-bit Python 3.4 is required; see "
            "xp\\README.md"
        )

    return win32com.client.Dispatch("Agent.Control.2")


class ClippyController(object):
    """Own one persistent Microsoft Agent character for this process."""

    def __init__(self, agent_factory=None):
        self._agent_factory = agent_factory or create_agent_control
        self._agent = None
        self._characters = None
        self._character = None
        self._animation_names = ()
        self._animation_name_set = set()

    @property
    def connected(self):
        return self._character is not None

    def connect(self):
        if self.connected:
            return

        agent = self._agent_factory()
        try:
            agent.Connected = True
            characters = agent.Characters
            characters.Load(CHARACTER_KEY, CHARACTER_PATH)
            character = characters.Character(CHARACTER_KEY)
            character.Balloon.Style = 7

            names = tuple(str(name) for name in character.AnimationNames)
            if not names:
                raise ControllerError(
                    "the installed Clippit character reported no animations"
                )
        except Exception:
            try:
                agent.Characters.Unload(CHARACTER_KEY)
            except Exception:
                pass
            raise

        self._agent = agent
        self._characters = characters
        self._character = character
        self._animation_names = names
        self._animation_name_set = set(names)

    def close(self):
        if self._character is not None:
            try:
                self._character.StopAll()
            except Exception:
                pass
        if self._characters is not None:
            try:
                self._characters.Unload(CHARACTER_KEY)
            except Exception:
                pass

        self._animation_names = ()
        self._animation_name_set = set()
        self._character = None
        self._characters = None
        self._agent = None

    def animation_names(self):
        self._require_connected()
        return list(self._animation_names)

    def show(self):
        self._require_connected()
        self._character.Show()

    def hide(self):
        self._require_connected()
        self._character.Hide()

    def move(self, x, y):
        self._require_connected()
        x = _coordinate("x", x)
        y = _coordinate("y", y)
        self._character.MoveTo(x, y)

    def speak(self, text):
        self._require_connected()
        self._character.Speak(_text(text))

    def think(self, text):
        self._require_connected()
        self._character.Think(_text(text))

    def play(self, animation_name):
        self._require_connected()
        if not isinstance(animation_name, str):
            raise InvalidParams("animation must be a string")
        if animation_name not in self._animation_name_set:
            raise InvalidParams(
                "animation is not present in the installed CLIPPIT.ACS: {0}".format(
                    animation_name
                )
            )
        self._character.Play(animation_name)

    def pump_messages(self):
        """Pump pending COM callbacks between synchronous controller steps."""
        try:
            import pythoncom
        except ImportError:
            return
        pythoncom.PumpWaitingMessages()

    def _require_connected(self):
        if not self.connected:
            raise ControllerError("ClippyController.connect() has not completed")


class JsonRpcServer(object):
    """A narrow JSON-RPC adapter; this is not yet an MCP server."""

    def __init__(self, controller):
        self.controller = controller
        self.should_stop = False

    def handle(self, request):
        if not isinstance(request, dict) or request.get("jsonrpc") != "2.0":
            raise InvalidRequest("request must be a JSON-RPC 2.0 object")

        if "id" in request:
            request_id = request["id"]
            invalid_boolean_id = isinstance(request_id, bool)
            valid_id_type = isinstance(
                request_id, (int, float, str, type(None))
            )
            if invalid_boolean_id or not valid_id_type:
                raise InvalidRequest("id must be a string, number, or null")

        method = request.get("method")
        if not isinstance(method, str):
            raise InvalidRequest("method must be a string")

        params = request.get("params", {})
        if not isinstance(params, dict):
            raise InvalidParams("params must be an object")

        if method == "clippy.animations":
            _exact_params(params, ())
            return {"animations": self.controller.animation_names()}
        if method == "clippy.show":
            _exact_params(params, ())
            self.controller.show()
            return {"queued": True}
        if method == "clippy.hide":
            _exact_params(params, ())
            self.controller.hide()
            return {"queued": True}
        if method == "clippy.move":
            _exact_params(params, ("x", "y"))
            self.controller.move(params["x"], params["y"])
            return {"queued": True, "x": params["x"], "y": params["y"]}
        if method == "clippy.speak":
            _exact_params(params, ("text",))
            self.controller.speak(params["text"])
            return {"queued": True}
        if method == "clippy.think":
            _exact_params(params, ("text",))
            self.controller.think(params["text"])
            return {"queued": True}
        if method == "clippy.play":
            _exact_params(params, ("animation",))
            self.controller.play(params["animation"])
            return {"queued": True, "animation": params["animation"]}
        if method == "clippy.shutdown":
            _exact_params(params, ())
            self.should_stop = True
            return {"closing": True}

        raise MethodNotFound("unknown method: {0}".format(method))


def _exact_params(params, expected):
    actual = set(params.keys())
    wanted = set(expected)
    if actual != wanted:
        raise InvalidParams(
            "expected params {0}; received {1}".format(
                sorted(wanted), sorted(actual)
            )
        )


def _coordinate(name, value):
    if isinstance(value, bool) or not isinstance(value, int):
        raise InvalidParams("{0} must be an integer".format(name))
    if value < MIN_COORDINATE or value > MAX_COORDINATE:
        raise InvalidParams(
            "{0} must be between {1} and {2}".format(
                name, MIN_COORDINATE, MAX_COORDINATE
            )
        )
    return value


def _text(value):
    if not isinstance(value, str):
        raise InvalidParams("text must be a string")
    if not value.strip():
        raise InvalidParams("text must not be empty")
    if len(value) > MAX_TEXT_LENGTH:
        raise InvalidParams(
            "text must be at most {0} characters".format(MAX_TEXT_LENGTH)
        )
    return value


def _error_response(request_id, code, message):
    return {
        "jsonrpc": "2.0",
        "id": request_id,
        "error": {"code": code, "message": message},
    }


def _write_json(stream, value):
    data = json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    stream.write(data.encode("utf-8") + b"\n")
    stream.flush()


def _log_internal_error(error):
    sys.stderr.write("Clippy controller error: {0}\n".format(repr(error)))
    sys.stderr.flush()


def serve(stdin=None, stdout=None, controller=None):
    input_stream = stdin or sys.stdin.buffer
    output_stream = stdout or sys.stdout.buffer
    active_controller = controller or ClippyController()
    server = JsonRpcServer(active_controller)

    try:
        active_controller.connect()
        for raw_line in input_stream:
            request = None
            request_id = None
            has_id = False
            try:
                request = json.loads(raw_line.decode("utf-8"))
                if isinstance(request, dict):
                    has_id = "id" in request
                    request_id = request.get("id")
                result = server.handle(request)
                if has_id:
                    _write_json(
                        output_stream,
                        {"jsonrpc": "2.0", "id": request_id, "result": result},
                    )
            except ValueError as error:
                _log_internal_error(error)
                _write_json(
                    output_stream, _error_response(None, -32700, "Parse error")
                )
            except InvalidRequest as error:
                _write_json(output_stream, _error_response(None, -32600, str(error)))
            except MethodNotFound as error:
                if has_id:
                    _write_json(
                        output_stream,
                        _error_response(request_id, -32601, str(error)),
                    )
            except InvalidParams as error:
                if has_id:
                    _write_json(
                        output_stream,
                        _error_response(request_id, -32602, str(error)),
                    )
            except Exception as error:
                _log_internal_error(error)
                if has_id:
                    _write_json(
                        output_stream,
                        _error_response(
                            request_id, -32603, "Microsoft Agent operation failed"
                        ),
                    )

            active_controller.pump_messages()
            if server.should_stop:
                break
    finally:
        active_controller.close()


if __name__ == "__main__":
    serve()
