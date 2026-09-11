# Windows XP Clippy controller

`clippy_agent.py` is a Python 3.4-compatible, process-persistent controller for
the real Microsoft Agent character installed with Office XP. It loads
`C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS` through
`Agent.Control.2`; it does not use Word's Office Assistant or screen
automation.

## XP dependency

The validated XP environment uses 32-bit Python 3.4.4 and pywin32 build 220.
That historical build is not available as a Python 3.4 wheel on current PyPI.
Install the official 32-bit Python 3.4 package from the archived pywin32
release:

```text
https://sourceforge.net/projects/pywin32/files/pywin32/Build%20220/
pywin32-220.win32-py3.4.exe
SHA-256 c86bea23fec5f353094b42ec0b48553db9152e6dfc00681df1545efd8af3c63b
```

Run the installer on XP's visible desktop. The legacy installer does not honor
quiet mode reliably when launched through SSH.

## Protocol foundation

Start `python -u xp\clippy_agent.py` and write one JSON-RPC 2.0 request per
UTF-8 line. The narrow method surface is:

```text
clippy.animations
clippy.show
clippy.hide
clippy.move       {"x": 650, "y": 420}
clippy.speak      {"text": "Hello"}
clippy.think      {"text": "Working on it"}
clippy.play       {"animation": "Greeting"}
clippy.shutdown
```

Character actions are asynchronous Microsoft Agent requests. A successful
JSON-RPC result means the request was queued without a synchronous COM error;
Phase 1 does not yet report Agent request completion. `clippy.play` accepts
only names enumerated from the loaded `CLIPPIT.ACS` during that process.
The installed `CheckingSomething`, `GetTechy`, `Searching`, `Thinking`, and
`Writing` animations repeat indefinitely, so the controller retains each
request object and stops that specific request after two seconds. This gives a
repeating animation visible runtime while allowing later queued actions to
advance; ordinary animations keep their native Agent-defined duration.
Batch arrays are deliberately rejected; the foundation accepts exactly one
request object per input line.

The phase 3 MCP server is the same Python process with the `--mcp` switch:

```text
python -u xp\clippy_agent.py --mcp
```

It pins MCP protocol version `2025-11-25`, uses the same one-message-per-line
UTF-8 stdio transport, and exposes only the fixed `clippy.*` tool set. The
server requires `initialize` followed by `notifications/initialized` before
listing or calling tools. Tool results include a short text content item and
structured content; Agent actions remain asynchronous and report
`queued: true`. EOF hides and unloads Clippy before the process exits.

The original JSON-RPC mode remains available for controller diagnostics. The
MCP framing is intentionally direct in Python so the XP process retains one
COM apartment and one persistent Agent connection; no shell, generic COM, or
 desktop-control surface is added.

## Persistent Codex connection

The visible XP session can keep a loopback MCP listener running with:

```text
python -u C:\\clippy\\xp\\clippy_agent.py --mcp-tcp 127.0.0.1 3211
```

Install `scripts\\service\\clippy_mcp_start.bat` in Brett's Startup folder so
the listener starts in the interactive desktop session. It binds only to XP
loopback. Codex reaches it through the XP-side `mcp_stdio_forward.py` bridge
over the existing SSH connection; the bridge carries MCP bytes but never
creates an Agent COM client.

On the Mac, register that bridge with the installed Codex CLI:

```text
codex mcp add clippy -- /usr/bin/ssh -T windows-xp "python -u C:\\clippy\\xp\\mcp_stdio_forward.py 127.0.0.1 3211"
```

The visible listener must already be running before Codex starts a Clippy MCP
session. The Codex client may need a new session after registration so the new
server enters its tool inventory.

The TCP service owns one visible `ClippyController` for its full process
lifetime and accepts up to 16 simultaneous SSH bridges. Each connection has
fresh MCP lifecycle state, while one FIFO lease allows only the earliest
eligible connection to call the six tools that change visible Clippy. Every
initialized connection can still list tools, call `clippy.animations`, inspect
its `clippy.queue_status`, or give up its lease/queue place with
`clippy.release`. A waiting visual call returns an immediate tool error with
its position and performs no COM action.

Active disconnect or release promotes the next waiter. When at least one user
is waiting, one minute without a visual tool call or five minutes of total
lease time stops outstanding Agent actions and rotates the active connection
to the queue tail. Timeouts do not interrupt a lone connection. A released
connection remains connected but idle and re-enters at the tail on its next
visual call.

Bridge EOF drains pending response bytes, closes only that client socket, and
then performs any lease handoff; it does not hide or unload the shared Agent
character. The controller is hidden and unloaded only when the visible TCP
service exits. This ordering prevents a slow Agent reset from leaving a
disconnected socket in `CLOSE_WAIT` and blocking the next Codex session.

On Codex stdin EOF, `mcp_stdio_forward.py` half-closes its TCP write side and
allows two seconds for the XP service to finish its final output. If the
service does not finish, the bridge fully closes the socket and exits instead
of waiting forever. The visible service uses one 50 ms `select` loop for all
clients and the Agent COM message pump. Client sockets are nonblocking and
each input/output buffer is capped at 64 KiB, so a slow client cannot block
other MCP sessions or Clippy animation progress.

## MCP diagnostics

The persistent service and each SSH bridge append separate JSONL diagnostics:

```text
C:\clippy\ClippyMcp.log
C:\clippy\ClippyMcpBridge.log
```

The service log covers listener startup, accept/session lifecycle, controller
connection, MCP method and tool dispatch, response writes, message pumping,
errors, and client close. The bridge log covers Codex stdin reads, TCP sends
and receives, stdout writes, EOF, errors, and bounded shutdown. Entries include
only routing metadata such as timestamp, PID, session/sequence number, method,
tool name, byte count, state, and error type. They never include request IDs,
params, tool arguments, message text, animation names supplied by a caller, or
full MCP payloads. Log writes are best-effort and cannot write to MCP stdout or
fail a request.

The bridge allows ten seconds to establish its loopback TCP connection, then
switches the connected socket back to blocking mode. This keeps its receive
thread alive across idle Codex periods while preserving the two-second bounded
shutdown after stdin closes.

Read the current logs over SSH without starting another COM controller:

```text
type C:\clippy\ClippyMcp.log
type C:\clippy\ClippyMcpBridge.log
```
