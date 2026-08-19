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
MCP_PROTOCOL_VERSION = "2025-11-25"
MCP_SERVER_INFO = {
    "name": "clippy-xp-agent",
    "version": "0.3.0",
}
MCP_NO_RESPONSE = object()


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
                self._character.Hide()
            except Exception:
                pass
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
    """The phase 1/2 private JSON-RPC adapter for controller diagnostics."""

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


class McpError(Exception):
    def __init__(self, code, message):
        Exception.__init__(self, message)
        self.code = code
        self.message = message


class McpServer(object):
    """Minimal MCP server for the fixed desktop-wide Clippy action surface."""

    STARTING = "STARTING"
    INITIALIZED = "INITIALIZED"
    READY = "READY"
    STOPPING = "STOPPING"

    def __init__(self, controller):
        self.controller = controller
        self.state = self.STARTING

    def handle(self, request):
        if not isinstance(request, dict) or request.get("jsonrpc") != "2.0":
            raise McpError(-32600, "Request must be a JSON-RPC 2.0 object")

        has_id = "id" in request
        if has_id:
            _validate_json_rpc_id(request["id"])

        method = request.get("method")
        if not isinstance(method, str):
            raise McpError(-32600, "Method must be a string")

        params = request.get("params", {})
        if not isinstance(params, dict):
            raise McpError(-32602, "Params must be an object")

        if method == "initialize":
            if not has_id:
                raise McpError(-32600, "initialize must be a request")
            return self._initialize(params)

        if self.state == self.STARTING:
            raise McpError(-32002, "Server has not been initialized")

        if method == "notifications/initialized":
            if has_id:
                raise McpError(-32600, "notifications/initialized is a notification")
            _exact_params(params, ())
            if self.state != self.INITIALIZED:
                raise McpError(-32002, "notifications/initialized is not expected")
            self.state = self.READY
            return MCP_NO_RESPONSE

        if self.state != self.READY:
            if not has_id:
                return MCP_NO_RESPONSE
            raise McpError(-32002, "Server is waiting for notifications/initialized")

        if method == "tools/list":
            _exact_params(params, ())
            return {"tools": _mcp_tools()}

        if method == "tools/call":
            _exact_params(params, ("name",), optional=("arguments",))
            if not isinstance(params.get("name"), str):
                raise McpError(-32602, "tools/call name must be a string")
            arguments = params.get("arguments", {})
            if not isinstance(arguments, dict):
                raise McpError(-32602, "tools/call arguments must be an object")
            return self._call_tool(params["name"], arguments)

        if not has_id:
            return MCP_NO_RESPONSE
        raise McpError(-32601, "Method not found")

    def _initialize(self, params):
        if self.state != self.STARTING:
            raise McpError(-32600, "initialize must be the first request")
        _exact_params(
            params,
            ("protocolVersion", "capabilities", "clientInfo"),
        )
        if params["protocolVersion"] != MCP_PROTOCOL_VERSION:
            raise McpError(
                -32602,
                "Unsupported protocol version; supported versions: {0}".format(
                    MCP_PROTOCOL_VERSION
                ),
            )
        if not isinstance(params["capabilities"], dict):
            raise McpError(-32602, "capabilities must be an object")
        client_info = params["clientInfo"]
        if not isinstance(client_info, dict):
            raise McpError(-32602, "clientInfo must be an object")
        if not isinstance(client_info.get("name"), str) or not isinstance(
            client_info.get("version"), str
        ):
            raise McpError(-32602, "clientInfo requires string name and version")

        try:
            self.controller.connect()
        except Exception:
            _log_internal_error("MCP controller initialization failed")
            raise McpError(-32603, "Clippy Agent initialization failed")

        self.state = self.INITIALIZED
        return {
            "protocolVersion": MCP_PROTOCOL_VERSION,
            "capabilities": {"tools": {}},
            "serverInfo": dict(MCP_SERVER_INFO),
        }

    def _call_tool(self, name, arguments):
        tools = _mcp_tool_names()
        if name not in tools:
            raise McpError(-32601, "Unknown tool")

        try:
            if name == "clippy.animations":
                _exact_params(arguments, ())
                animations = self.controller.animation_names()
                return _tool_success(
                    "Clippy reported {0} installed animations.".format(
                        len(animations)
                    ),
                    {"animations": animations},
                )
            if name == "clippy.show":
                _exact_params(arguments, ())
                self.controller.show()
                return _tool_success(
                    "Clippy show request queued.", {"queued": True}
                )
            if name == "clippy.hide":
                _exact_params(arguments, ())
                self.controller.hide()
                return _tool_success(
                    "Clippy hide request queued.", {"queued": True}
                )
            if name == "clippy.move":
                _exact_params(arguments, ("x", "y"))
                self.controller.move(arguments["x"], arguments["y"])
                return _tool_success(
                    "Clippy move request queued.",
                    {"queued": True, "x": arguments["x"], "y": arguments["y"]},
                )
            if name == "clippy.speak":
                _exact_params(arguments, ("text",))
                self.controller.speak(arguments["text"])
                return _tool_success(
                    "Clippy speak request queued.", {"queued": True}
                )
            if name == "clippy.think":
                _exact_params(arguments, ("text",))
                self.controller.think(arguments["text"])
                return _tool_success(
                    "Clippy think request queued.", {"queued": True}
                )
            if name == "clippy.play":
                _exact_params(arguments, ("animation",))
                self.controller.play(arguments["animation"])
                return _tool_success(
                    "Clippy animation request queued.",
                    {"queued": True, "animation": arguments["animation"]},
                )
        except InvalidParams as error:
            return _tool_error(_safe_tool_param_error(error))
        except Exception:
            _log_internal_error("MCP Clippy tool operation failed")
            return _tool_error("Clippy could not queue that request.")

        raise McpError(-32601, "Unknown tool")

    def close(self):
        self.state = self.STOPPING


def _validate_json_rpc_id(request_id):
    if isinstance(request_id, bool) or not isinstance(
        request_id, (int, float, str, type(None))
    ):
        raise McpError(-32600, "id must be a string, number, or null")


def _mcp_tools():
    return [
        {
            "name": "clippy.animations",
            "description": (
                "List the exact animation names installed in Clippit's "
                "CLIPPIT.ACS character."
            ),
            "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "name": "clippy.show",
            "description": "Queue an asynchronous request to show Clippy.",
            "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "name": "clippy.hide",
            "description": "Queue an asynchronous request to hide Clippy.",
            "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        },
        {
            "name": "clippy.move",
            "description": "Queue an asynchronous request to move Clippy on screen.",
            "inputSchema": {
                "type": "object",
                "properties": {
                    "x": {"type": "integer", "minimum": MIN_COORDINATE, "maximum": MAX_COORDINATE},
                    "y": {"type": "integer", "minimum": MIN_COORDINATE, "maximum": MAX_COORDINATE},
                },
                "required": ["x", "y"],
                "additionalProperties": False,
            },
        },
        {
            "name": "clippy.speak",
            "description": (
                "Queue an asynchronous speech request; text must be 1-2000 "
                "characters and is shown by the real Agent character."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "text": {"type": "string", "minLength": 1, "maxLength": MAX_TEXT_LENGTH},
                },
                "required": ["text"],
                "additionalProperties": False,
            },
        },
        {
            "name": "clippy.think",
            "description": (
                "Queue an asynchronous thought request; text must be 1-2000 "
                "characters and is shown by the real Agent character."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {
                    "text": {"type": "string", "minLength": 1, "maxLength": MAX_TEXT_LENGTH},
                },
                "required": ["text"],
                "additionalProperties": False,
            },
        },
        {
            "name": "clippy.play",
            "description": (
                "Queue an asynchronous installed animation by exact name; "
                "call clippy.animations first to discover valid names."
            ),
            "inputSchema": {
                "type": "object",
                "properties": {"animation": {"type": "string", "minLength": 1}},
                "required": ["animation"],
                "additionalProperties": False,
            },
        },
    ]


def _mcp_tool_names():
    return set(tool["name"] for tool in _mcp_tools())


def _tool_success(summary, structured_content):
    return {
        "content": [{"type": "text", "text": summary}],
        "structuredContent": structured_content,
    }


def _tool_error(message):
    return {
        "content": [{"type": "text", "text": message}],
        "structuredContent": {"error": message},
        "isError": True,
    }


def _safe_tool_param_error(error):
    message = str(error)
    if "animation" in message:
        if "not present" in message:
            return "The requested animation is not installed in Clippit."
        return "animation must be a string"
    if "text" in message:
        return "text must be a non-empty string of at most 2000 characters"
    if message.startswith("x ") or message.startswith("y "):
        return "x and y must be integers between -32768 and 32767"
    if message.startswith("expected params"):
        return "Invalid tool arguments"
    return "Invalid tool arguments"


def _exact_params(params, expected, optional=()):
    actual = set(params.keys())
    wanted = set(expected)
    allowed = wanted.union(set(optional))
    if not actual.issubset(allowed) or not wanted.issubset(actual):
        raise InvalidParams(
            "expected params {0}; received {1}".format(
                sorted(wanted), sorted(actual)
            )
        )


def _mcp_error_response(request_id, error):
    return {
        "jsonrpc": "2.0",
        "id": request_id,
        "error": {"code": error.code, "message": error.message},
    }


def serve_mcp(stdin=None, stdout=None, controller=None):
    """Serve MCP over one UTF-8 JSON-RPC message per stdin line."""
    input_stream = stdin if stdin is not None else sys.stdin.buffer
    output_stream = stdout if stdout is not None else sys.stdout.buffer
    active_controller = controller or ClippyController()
    server = McpServer(active_controller)

    try:
        for raw_line in input_stream:
            request = None
            request_id = None
            has_id = False
            is_notification = False
            try:
                try:
                    request = json.loads(raw_line.decode("utf-8"))
                except (UnicodeDecodeError, ValueError):
                    raise McpError(-32700, "Parse error")

                if isinstance(request, dict):
                    has_id = "id" in request
                    is_notification = not has_id and "method" in request
                    candidate_id = request.get("id")
                    if not isinstance(candidate_id, bool) and isinstance(
                        candidate_id, (int, float, str, type(None))
                    ):
                        request_id = candidate_id
                result = server.handle(request)
                if has_id and result is not MCP_NO_RESPONSE:
                    _write_json(
                        output_stream,
                        {"jsonrpc": "2.0", "id": request_id, "result": result},
                    )
            except McpError as error:
                if not is_notification:
                    _write_json(output_stream, _mcp_error_response(request_id, error))
            except Exception:
                _log_internal_error("Unhandled MCP server error")
                if has_id and not is_notification:
                    _write_json(
                        output_stream,
                        _mcp_error_response(
                            request_id,
                            McpError(-32603, "Internal MCP server error"),
                        ),
                    )

            active_controller.pump_messages()
            if server.state == server.STOPPING:
                break
    finally:
        server.close()
        active_controller.close()


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
    if "--mcp" in sys.argv[1:]:
        serve_mcp()
    else:
        serve()
