# Clippy MCP design

Status: phase 3 implemented with concurrent FIFO admission, 2026-09-10

This specification defines the next protocol work after the completed Phase
1/2 global Microsoft Agent foundation. It is intentionally Clippy-specific:
the real Windows XP Microsoft Agent character is the primary interface, and
the Word Office Assistant remains a later, separate capability.

## Scope and decisions

The desktop-wide path is authoritative for this design:

```text
MCP client / modern agent
             │ stdio JSON-RPC
             ▼
XP Clippy MCP process or adapter
             │ one persistent COM client
             ▼
Agent.Control.2 → CLIPPIT.ACS → visible Clippy
```

The existing `xp/clippy_agent.py` is the Phase 1/2 implementation baseline.
It owns one `Agent.Control.2` connection and one loaded character, validates
all parameters, and only plays animation names enumerated from:

```text
C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS
```

### Verified XP constraints

These constraints are part of the runtime contract, not setup suggestions:

* The controller must remain Python 3.4-compatible and runs on the verified
  32-bit Python 3.4.4 installation.
* The working COM binding is 32-bit pywin32 build 220. The archived installer
  is documented in `xp/README.md`.
* The Agent process must launch in the visible XP desktop session. An Agent
  client launched through SSH can be non-visible or hang while setting
  `Connected = True`; SSH is for transfer, fake-COM tests, and diagnostics.
* `comtypes==1.2.1` was tested as a Python 3.4-compatible alternative, but
  both dynamic and generated dispatch hung at `Connected=True` on this VM. It
  is not the controller dependency.
* `clippy.play` must guard against the exact runtime `AnimationNames` returned
  by the installed `CLIPPIT.ACS`; no animation list copied from documentation
  or another character is authoritative.
* No MCP tool may expose arbitrary shell execution, unrestricted COM object
  creation, keyboard/mouse injection, registry mutation, or generic desktop
  control.

The following are outside this design:

* arbitrary shell execution, unrestricted COM creation, input injection, or
  generic window manipulation;
* moving modern agent logic or an MCP runtime into the XP guest beyond the
  smallest required controller process;
* treating Word's Office Assistant as the global host;
* hidden or simulated Word UI interaction.

## Phased plan

| Phase | Status | Deliverable | Exit evidence |
| --- | --- | --- | --- |
| 1/2 | Complete | Persistent Python controller with show, hide, move, speak, think, runtime animation enumeration/guarding, and line JSON-RPC foundation. | XP Python 3.4.4 tests pass; visible UTM session shows Clippy, a verified animation, Think/Speak, invalid-animation rejection, and clean shutdown. |
| 3 | Complete | Minimal MCP lifecycle, stdio framing, `tools/list`, initial Clippy tools, and FIFO admission for concurrent persistent connections. | Python fixture tests cover initialize, capability negotiation, tool discovery/calls, notifications, malformed requests, unsupported versions, tool errors, EOF shutdown, queue ordering and rotation, disconnect promotion, and nonblocking partial writes. |
| 4 | Planned | Tightly scoped read-only XP automation with an explicit confirmation gate. | Every read-only tool has an allowlist, bounded output, confirmation behavior, and a denied-path test. |
| 5 | Planned | Word-native Office Assistant integration as a separate capability with intentional visible handoff. | User-visible XP acceptance shows the original Assistant UI; no hidden Word automation is used. |
| 6 | Planned | Full XP validation and documentation synchronization. | Reproducible test commands, fresh screenshots/logs, architecture/devlog updates, and no generated artifacts committed. |

Phase 3 supports MCP protocol versions `2025-06-18` and `2025-11-25`. The
server returns the client-requested supported version in its initialize result;
`2025-11-25` remains the implementation's preferred/current version. This
compatibility is required by the Codex desktop app-server, which currently
initializes stdio servers with `2025-06-18`.

## Phase 3: minimal MCP server

### Transport and lifecycle

Use one UTF-8 JSON-RPC message per stdin line. A message must not contain an
embedded newline. stdout is reserved for valid MCP messages; diagnostics go to
stderr. The server must not emit a startup banner.

The server state machine is:

1. `STARTING`: create or attach to the persistent controller, but do not
   expose tools before protocol initialization.
2. `INITIALIZED`: accept only `initialize` first, negotiate a supported protocol
   version, and return implementation information plus the `tools` capability.
3. `READY`: require the `notifications/initialized` notification before
   serving `tools/list` or `tools/call`.
4. `STOPPING`: close stdin/EOF, stop outstanding Agent requests where
   possible, hide the character, unload `Clippy`, and exit cleanly.

For direct stdio mode, EOF still owns the whole process lifecycle and performs
the `STOPPING` cleanup above. For persistent `--mcp-tcp` mode, each accepted
socket has an independent state machine and bridge EOF ends only that MCP
protocol session. The service drains any buffered final response, closes the
socket, discards that session state, and then removes its FIFO lease entry. The
one visible Agent controller remains loaded for other clients and is hidden,
stopped, and unloaded only when the listener process exits. Closing the socket
before lease handoff prevents a slow Agent reset from retaining a dead client
socket in `CLOSE_WAIT`.

The minimal server does not advertise prompts, resources, sampling, roots,
elicitation, tasks, logging, or `listChanged`. It must reject requests that
depend on an unadvertised capability. Batch messages are not part of this
initial one-message-per-line contract.

Required lifecycle messages:

```json
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"<supported-version>","capabilities":{},"clientInfo":{"name":"<client>","version":"<client-version>"}}}
{"jsonrpc":"2.0","method":"notifications/initialized"}
```

The initialize result must contain the negotiated `protocolVersion`,
`capabilities: {"tools": {}}`, and `serverInfo` identifying this as the
Clippy XP Agent server. An unsupported version is a JSON-RPC `-32602` error
with a safe supported-version list; it must not start a second controller.
Standard MCP request metadata under `params._meta` is accepted on lifecycle
and tool requests; it is transport metadata and is not passed into tool
argument validation.
Normal stdio shutdown is initiated by closing stdin and waiting for the server
to exit. A private `clippy.shutdown` method may remain as a test-only adapter
operation, but it is not an MCP tool.

### Initial tools

`tools/list` returns exactly these tools, with JSON Schema `inputSchema`
objects and descriptions that state that Agent actions are asynchronous:

| Tool | Arguments | Result contract |
| --- | --- | --- |
| `clippy.animations` | none | `structuredContent.animations`: exact runtime names from `CLIPPIT.ACS`. |
| `clippy.show` | none | `structuredContent.queued: true`. |
| `clippy.hide` | none | `structuredContent.queued: true`. |
| `clippy.move` | required integer `x`, `y`; each from -32768 through 32767 | `structuredContent: {"queued":true,"x":x,"y":y}`. |
| `clippy.speak` | required non-empty string `text`, max 2000 characters | `structuredContent.queued: true`. |
| `clippy.think` | required non-empty string `text`, max 2000 characters | `structuredContent.queued: true`. |
| `clippy.play` | required string `animation` | Queue only if the exact name was returned by `clippy.animations`; otherwise a tool error and no COM call. Known repeating animations are stopped after two seconds so later queued actions can advance. |
| `clippy.queue_status` | none | Persistent TCP mode only. Reports this connection as `active`, `waiting`, or `idle`, including FIFO position and timeout details where applicable. |
| `clippy.release` | none | Persistent TCP mode only. Removes this connection's active lease or waiting place and promotes the next waiter when needed. |

`tools/call` receives the tool name and an object of arguments. A successful
call returns a `content` array containing a short text summary and a matching
`structuredContent` object. Agent request completion is not implied by
`queued: true`; until completion/error observation exists, the tool description
and result must use that wording.

The six tools that change the visible character require the connection to hold
the persistent service's lease. A waiting call returns a normal MCP tool result
with `isError: true`, its current FIFO `position`, and `waitingCount`; it must
not invoke COM. `clippy.animations` and both queue tools remain available to
every initialized connection. Direct `--mcp` stdio mode has no shared lease and
continues to advertise only the original seven tools.

Microsoft Agent serializes its asynchronous character requests. The installed
`CheckingSomething`, `GetTechy`, `Searching`, `Thinking`, and `Writing`
animations repeat rather than completing on their own. The controller retains
the request object returned for each of those exact runtime animation names;
its existing 50 ms COM message pump stops that specific request after two
seconds. This keeps the animation visibly active for a bounded interval while
preserving any speech, movement, visibility, or other animation already queued
behind it. Non-repeating animations retain their native duration and Agent
queue order.

Tool failures that are part of normal operation—bad arguments, an animation
not present in the installed character, or a synchronous Agent failure—return a
successful JSON-RPC response with `isError: true` and a safe text/structured
error. Unknown tools, malformed MCP requests, lifecycle violations, and
unsupported protocol versions return JSON-RPC error responses. Never include
filesystem paths, COM exception details, or arbitrary command text in a user
visible error.

## Phase 4: read-only XP automation

Phase 4 is an allowlisted observation layer, not a general desktop API. The
first implementation may expose only data needed to explain Clippy state, such
as the controller connection state, loaded character identity, installed
animation names, and a bounded active-window summary. It must not read files,
clipboard contents, environment secrets, arbitrary window text, or process
memory without a separate approved tool and review.

Every new read-only tool must declare:

* the exact Windows API/query and fields returned;
* a hard output size and timeout;
* why the data is needed by Clippy;
* whether it requires explicit user confirmation;
* a denial path that performs no XP call.

The MCP host must present the exact tool name, arguments, and intended read to
the user before any confirmation-gated call. “Read-only” does not authorize a
broader inspection surface by implication. There is no Phase 4 write,
keyboard/mouse injection, shell, registry mutation, or unrestricted COM tool.

## Phase 5: Word-native Assistant integration

Word support remains a separate Office-specific capability layered beside the
global Agent host. It may use the existing native `ClippyShim.dll`, Word's
`Assistant.NewBalloon`, and the authentic Assistant question editor. It must
not replace the global controller or make Word a prerequisite for desktop-wide
Clippy.

Any operation that focuses Word, opens the Assistant, submits a question, or
changes Word content requires an intentional visible handoff: the user must
see which Word instance is targeted and approve the exact action before it is
performed. SSH may copy files and inspect logs, but it is not evidence of
visible Word acceptance. UTM Capture Input remains prohibited.

## Phase 6: validation and documentation

Each implementation milestone must update this specification, the architecture
source of truth, and the development log before its commit. Validation must
include:

* host-side protocol and schema tests with no XP dependency;
* the same controller tests under XP's Python 3.4.4 interpreter;
* visible UTM acceptance from the XP desktop, with fresh screenshots for
  character presence and a balloon/later animation;
* confirmation that no stale `python.exe`, `agentsvr.exe`, Word demo process,
  generated binary, or runtime log is accidentally committed;
* explicit recording of limitations, especially queued Agent requests and
  visible-session lifecycle behavior.

## Resolved architecture decision

The two placements considered for the MCP framing were:

1. **Direct MCP in Python** — extend `xp/clippy_agent.py` with MCP lifecycle,
   schema, and tool dispatch. This keeps one process and one COM apartment, but
   couples the Python 3.4 compatibility surface to MCP protocol evolution.
2. **MCP adapter around the current controller** — keep the tested line
   JSON-RPC controller as a narrow child process and put MCP lifecycle/tool
   translation in a modern host process. This isolates XP compatibility and
   makes protocol evolution easier, but adds process supervision and a second
   error boundary.

Phase 3 selects **direct MCP in Python**. `xp/clippy_agent.py --mcp` adds the
MCP lifecycle and tool translation around the existing `ClippyController`,
keeping one process, one COM apartment, and one persistent Agent connection.
The original line JSON-RPC mode remains available as a private diagnostic
adapter operation. The XP-facing action vocabulary and safety constraints in
this document are unchanged, and Agent calls still report `queued: true`
until completion/error observation is implemented.

### Persistent Codex service boundary

For a persistent Codex connection, run the direct MCP server in the visible XP
desktop with `--mcp-tcp 127.0.0.1 3211`. The listener is loopback-only and is
started from the interactive user's Startup folder. Codex launches
`mcp_stdio_forward.py` over the existing `windows-xp` SSH alias; that bridge
forwards MCP stdin/stdout to the XP loopback listener and never creates COM.
This preserves the visible-session requirement while allowing Codex's stdio
MCP client to register the service with `codex mcp add clippy`.

The TCP listener accepts at most 16 simultaneous clients and reuses one visible
`ClippyController` across them. A single COM-owner-thread `select` loop gives
each socket independent protocol state and nonblocking 64 KiB input/output
buffers. Accept order defines the FIFO lease: the first client is active and
later clients wait. Active disconnect or `clippy.release` promotes the next
waiter. If anyone is waiting, one minute without an active visual-tool attempt
or five minutes total rotates the active client to the tail after `StopAll`;
the time limits never interrupt a lone client. A released connection stays
connected but idle, and its next visual action re-enters at the queue tail.

Client EOF drains buffered output and fully closes that socket before the
lease manager resets Clippy and promotes another session. The stdio bridge
half-closes its write side on Codex EOF, waits up to two seconds for remaining
server output, then fully closes the socket if the server has not finished.
This bounds bridge shutdown, prevents closed sessions from indefinitely
occupying the listener backlog, and lets a slow bridge coexist without
blocking other clients or Agent message pumping.

Protocol references consulted for this draft:

* [MCP lifecycle](https://modelcontextprotocol.io/specification/2025-11-25/basic/lifecycle)
* [MCP stdio transport](https://modelcontextprotocol.io/specification/2025-11-25/basic/transports)
* [MCP tools](https://modelcontextprotocol.io/specification/2025-11-25/server/tools)
