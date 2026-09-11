Clippy Possession — Project One-Pager

Goal

Resurrect the original Microsoft Office Clippy and turn him into a general-purpose AI assistant without replacing the original UI.

Rather than recreating Clippy as a modern app or animation, the project uses the real Microsoft Agent / Office Assistant stack running under Windows XP and Office XP. A modern AI agent supplies the intelligence while the vintage Office Assistant remains the visible interface.

The end state is an authentic Clippy that can accept natural-language requests, reason with a modern LLM, use tools on the host computer, and respond using his original speech balloons, animations, questions, and controls.

⸻

Core Principle

Keep the two eras separate:

* Windows XP owns Clippy.
* The modern host owns intelligence and tools.

XP should contain as little modern infrastructure as possible. It acts primarily as Clippy’s body and compatibility environment.

Current Status — 2026-09-10

The original Word integration's first four milestones and first AI response
checkpoint remain complete:

* native code can control the installed Clippit character
* the authentic Office Assistant query editor and Search command are identified
* `ClippyShim.dll` intercepts Search and unmodified Enter inside Word
* XP posts the captured text to the macOS TypeScript server
* Clippy loops his authentic Thinking animation while that request is pending
* the server sends the text to OpenAI's Responses API with `gpt-5.6-luna`
* Word renders model-generated Markdown-lite with native Assistant balloon
  colors, underlining, and list labels

The first foundation for a desktop-wide host is complete. A Python
3.4-compatible process on XP owns the real Microsoft Agent character
independently of Word and exposes a fixed JSON-RPC command surface for show,
hide, move, speak, think, installed-animation enumeration, and guarded play.
The same process now exposes phase 3 MCP over stdio with `--mcp`, pinned to
protocol version `2025-11-25`, while keeping one `Agent.Control.2` connection
and one loaded `CLIPPIT.ACS` character for the full session.

Microsoft Agent is now the primary desktop-wide Clippy host. The Word COM
add-in remains a working, Office-specific input and rich-balloon integration;
it is not the component boundary for global Clippy control. Direct MCP in the
Python process is the selected phase 3 boundary: it preserves the single XP
COM apartment and keeps the existing JSON-RPC mode available for diagnostics.
MCP tools and lifecycle negotiation are implemented; Agent request completion
remains queued-only until completion/error observation is added.

The persistent TCP boundary now multiplexes up to 16 simultaneous MCP
connections while keeping one visible `Agent.Control.2` controller on its COM
owner thread. Each socket has independent MCP lifecycle state plus bounded
64 KiB input and output buffers. The first accepted connection receives the
Clippy lease and later connections wait in FIFO accept order; all connections
can initialize, list tools, and inspect their queue position immediately.
Visual tool calls from a waiting connection return a tool error with its queue
position and perform no Agent call.

An active connection keeps its lease for at most five minutes, or until it has
been inactive for one minute, while another connection is waiting. Either
timeout rotates it to the queue tail. Explicit `clippy.release`, disconnect,
or a timeout stops outstanding Agent actions and promotes the next connection;
a lone active connection is not interrupted by a timeout. Bridge EOF remains
a protocol-session boundary rather than an Agent lifetime boundary: the dead
client socket closes before lease handoff, and the one visible controller stays
loaded until the service process exits. The SSH bridge retains its bounded
two-second EOF close path and uses its ten-second timeout only while opening
the TCP connection.

The persistent TCP owner keeps the Agent apartment responsive even when MCP
traffic is idle. A single `select` loop polls the listener and every readable
or writable client socket in 50 ms intervals, pumps
`pythoncom.PumpWaitingMessages()` on the COM-owning main thread, and flushes
responses incrementally. Nonblocking per-client output prevents a slow bridge
from stalling other users or the COM message pump. Direct `--mcp` stdio mode
remains request-driven and retains the original seven-tool behavior; the
deployed Codex path uses the continuously pumped `--mcp-tcp` service and adds
`clippy.queue_status` and `clippy.release`.

The boundary now records metadata-only JSONL diagnostics on XP. The persistent
service appends to `C:\clippy\ClippyMcp.log`; each SSH stdio bridge appends to
`C:\clippy\ClippyMcpBridge.log`. Entries correlate process, TCP session,
request sequence, MCP method/tool name, byte count, lifecycle state, response
write, and error type. They deliberately exclude request IDs, params, tool
arguments, message text, animation arguments, and complete protocol payloads.
Logging is best-effort and never writes to MCP stdout.

The initial eve agent under `clippy-agent/` now consumes this same boundary.
Eve connections require Streamable HTTP or SSE rather than a spawned stdio
server, so `mcp-proxy` on macOS sits between eve and the existing SSH bridge.
`agent/connections/clippy.ts` registers the fixed Clippy tool allowlist at the
ngrok HTTPS endpoint by default. ngrok forwards to the proxy's loopback-only
listener at `http://127.0.0.1:3212/mcp`; `npm run clippy:mcp` launches the
existing `mcp_stdio_forward.py` command over the `windows-xp` SSH alias. The
proxy accepts modern `2026-07-28` Streamable HTTP from eve while independently
negotiating the XP server's legacy MCP version upstream. A different endpoint,
including the direct loopback URL for local-only use, may be supplied through
`CLIPPY_MCP_URL`, with an optional `X-API-Key` sourced from
`CLIPPY_MCP_API_KEY`. The public tunnel must be authenticated with
`MCP_PROXY_API_KEY` before it is treated as a durable deployment boundary.

┌──────────────────────────── macOS ────────────────────────────┐
│                                                              │
│  Modern AI Agent                                             │
│  ├── LLM / reasoning                                         │
│  ├── conversation + memory                                   │
│  ├── filesystem                                              │
│  ├── shell / coding tools                                    │
│  ├── browser / APIs                                          │
│  └── other MCP/tool integrations                             │
│                 │                                            │
│                 │ TCP / HTTP / simple RPC                    │
└─────────────────┼────────────────────────────────────────────┘
                  │
                  │ UTM virtual network
                  ▼
┌──────────────────── Windows XP SP3 x86 ──────────────────────┐
│                                                              │
│  Microsoft Agent (primary desktop-wide character host)       │
│      │                                                       │
│      └── xp/clippy_agent.py                                  │
│             ├── persistent Agent.Control.2 client            │
│             ├── installed-animation guard                    │
│             └── narrow stdin JSON-RPC actions                │
│                                                              │
│  Office XP (working Office-specific path)                     │
│      └── ClippyShim.dll                                      │
│             ├── capture Word Assistant input                 │
│             ├── send requests to Mac                         │
│             └── render native rich Assistant balloons        │
│                                                              │
│  Development / control infrastructure                        │
│  ├── Visual Studio / native C++ toolchain                    │
│  ├── SSH                                                     │
│  ├── SCP                                                     │
│  └── WindowWatch / debugging utilities                       │
└──────────────────────────────────────────────────────────────┘

Windows-Side Architecture

Microsoft Agent + Office Assistant

The authentic Microsoft stack remains responsible for:

* rendering Clippit.acs
* animations
* movement
* speech balloons
* Office Assistant interactions
* selectable questions/buttons
* the original “What would you like to do?” interface

Microsoft Agent itself supplies the animated character globally, while Office
adds richer interactive UI within Office applications. New desktop-wide work
must use Microsoft Agent as its primary character boundary and must not require
Word to be running. The Office Assistant path remains available when an
Office-native question editor or rich Office balloon is specifically needed.

Python Clippy controller

`xp/clippy_agent.py` is the Phase 1/2 global-host foundation. It runs under the
installed 32-bit Python 3.4.4 and uses pywin32 build 220 to drive the real
`Agent.Control.2` COM object. It always loads:

```text
C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS
```

The controller owns that COM object and character until stdin closes or
`clippy.shutdown` is received. Its JSON-RPC 2.0 methods are limited to:

```text
clippy.animations
clippy.show
clippy.hide
clippy.move
clippy.speak
clippy.think
clippy.play
clippy.shutdown
```

`clippy.play` accepts an animation only when the exact name was enumerated
from the loaded character during that process. Text and coordinates are
validated, extra parameters are rejected, and the adapter contains no shell,
filesystem, window-input, or generic COM invocation feature. Microsoft Agent
actions are asynchronous; a `{"queued":true}` result means the COM call
returned without a synchronous error, not that the animation request later
completed successfully.

With `python -u xp\clippy_agent.py --mcp`, the same controller serves MCP
protocol version `2025-11-25` over newline-delimited UTF-8 stdio. It requires
`initialize` and `notifications/initialized`, then exposes exactly the seven
`clippy.*` tools through `tools/list` and `tools/call`. Tool actions return
short text content plus structured content, with `queued: true` for
asynchronous Agent requests. The MCP path has no prompts, resources, MCP
logging capability, sampling, shell, input injection, generic COM, or broad
desktop-control capability. Its local metadata-only diagnostic files are an
implementation debugging aid, not a source of request content. EOF hides and
unloads Clippy.

Persistent Codex service boundary

Codex runs on macOS, while the real Agent COM client must remain in the visible
XP desktop. The persistent integration therefore has three narrow pieces:

```text
Codex MCP stdio
      │ SSH stdin/stdout
      ▼
XP mcp_stdio_forward.py ── TCP loopback ──> visible clippy_agent.py --mcp-tcp
                                                    │
                                                    ▼
                                             Agent.Control.2
```

The XP listener binds only to `127.0.0.1:3211` and is started by the user's
interactive Startup entry from `scripts/service/clippy_mcp_start.bat`. The
SSH bridge carries protocol bytes only and has no COM, shell-tool, or desktop
control surface. A Codex MCP registration starts one bridge per session. The
visible service accepts up to 16 bridges concurrently and assigns each an
independent MCP state machine. One FIFO lease gates the six visual action
tools; the active connection can act while waiting connections receive
immediate position-aware tool errors. `clippy.queue_status` reports `active`,
`waiting`, or `idle`, and `clippy.release` voluntarily gives up either the
active lease or a waiting place. One-minute inactivity and five-minute
absolute limits rotate the active lease only when someone is waiting.

On bridge EOF, the service drains any already-buffered response, closes that
client socket, and then performs lease handoff; the visible Agent controller
remains loaded for the process lifetime. The bridge waits at most two seconds
for final server output before forcing its local socket closed. One 50 ms
`select` loop handles listener accept, nonblocking reads and writes, lease
expiry, and Agent COM pumping on the same owner thread. Per-session input and
output are each capped at 64 KiB so one incomplete or slow connection cannot
consume unbounded memory or block another client.

The two XP JSONL logs make the transport boundary observable without copying
MCP contents. Bridge events show stdin read, TCP send, TCP receive, stdout
write, EOF, and bounded close. Service events show accept, controller connect,
method dispatch, tool dispatch, response write, message pump, and socket close.
The absence of a matching event therefore identifies which side of the
Codex-to-XP path stopped making progress while preserving user text.

The XP dependency is intentionally pinned to the period-compatible 32-bit
pywin32 build 220 installer. `comtypes` 1.2.1 imports on Python 3.4 and can
create `Agent.Control.2`, but both dynamic and generated dispatch hung when
setting `Connected = True` in this environment. The validated controller uses
pywin32 and must be launched from XP's visible desktop; an SSH-launched Agent
client can attach to a non-visible desktop or hang during connection.

Eve client adapter

The eve application in `clippy-agent/` uses the same XP service without adding
another COM owner or another XP protocol. Its development path is:

```text
eve connection_search / connection tool
             │ modern Streamable HTTP, HTTPS
             ▼
      ngrok public tunnel
             │ HTTP, 127.0.0.1:3212
             ▼
       macOS mcp-proxy
             │ child-process stdio
             ▼
  ssh windows-xp mcp_stdio_forward.py
             │ XP loopback TCP, 127.0.0.1:3211
             ▼
 visible clippy_agent.py --mcp-tcp
```

The connection allowlist repeats all nine tools published by the persistent
service: animation enumeration, queue status and release, plus show, hide,
move, speak, think, and guarded play. The HTTP proxy deliberately keeps its
downstream and upstream protocol negotiations separate: eve can use modern
`2026-07-28` request/response POSTs even though the XP server currently
negotiates `2025-06-18` or `2025-11-25`. This avoids the legacy session GET
whose response headers can remain buffered by an HTTP/2 tunnel. The proxy
multiplexes eve's HTTP requests onto one upstream stdio bridge, so the XP lease
manager sees the proxy as one MCP connection; direct Codex bridges and other XP
clients still participate in the service's normal FIFO admission.

ClippyShim.dll

A native Win32/x86 component loaded into Office is the Office-specific XP
integration layer. It is no longer the primary boundary for desktop-wide
Clippy control.

Responsibilities:

Office input
    ↓
ClippyShim.dll
    ↓
Mac AI agent
    ↓
ClippyShim.dll
    ↓
Office Assistant / Microsoft Agent
    ↓
📎 response

The shim should remain intentionally thin. It should not contain the LLM or general-purpose agent logic.

WindowWatch.exe

A passive reverse-engineering utility used to identify the implementation of Office’s original Assistant input UI.

It:

* runs invisibly
* never steals focus
* captures the desktop window hierarchy with F11
* creates sequential window dumps
* records HWNDs, classes, parents, process IDs, styles, focus, etc.

The initial before/after comparison of “What would you like to do?” is
complete and identified Word's `RichEdit20W` control ID `6` as the query
editor. WindowWatch remains useful for validating that hierarchy and measuring
subsequent UI changes without disturbing the visible XP desktop.

⸻

Development Architecture

A modern coding agent runs on the Mac but treats XP as a remote development machine.

Coding Agent on Mac
       │
       ├── SSH → execute commands
       ├── SCP → transfer source/binaries
       └── SSH → read build output
                    │
                    ▼
               Windows XP
                    │
               build.bat
                    │
              C++ compiler
                    │
              Clippy code

This lets the coding agent autonomously:

* inspect the XP environment
* write source code
* create build scripts
* compile Win32/x86 projects
* diagnose compiler errors
* run experiments
* inspect logs
* iterate against the actual Office XP environment

No modern coding-agent runtime needs to run on XP itself.

⸻

Major Milestones

Global Microsoft Agent host track — Phase 3 MCP complete

The persistent phase 3 boundary gained concurrent FIFO admission on
2026-09-10. One event loop now serves up to 16 independently initialized MCP
connections while the first accepted connection alone holds the visible
Clippy lease. Waiting callers can inspect their position or release it, and
active disconnect, explicit release, one-minute inactivity, or the five-minute
absolute limit promotes the next connection after stopping outstanding Agent
actions. The time limits rotate only when another connection is waiting.
Bounded nonblocking per-session buffers keep a slow client from stalling other
users or the single COM owner thread. Host and XP Python 3.4 tests cover FIFO
ordering, release/requeue, both timeout paths, disconnect promotion, partial
writes, and EOF response draining. Two live SSH bridges initialized
simultaneously; the waiting bridge was denied a visible action, then was
promoted after the active bridge released Clippy and moved the real character
on the UTM XP desktop.

The desktop-wide track now has a persistent Python controller and a narrow,
tested newline JSON-RPC adapter. The real XP acceptance covered runtime
animation enumeration, show, hide, move, Greeting play, think, speak, rejection
of a fabricated animation, and clean shutdown. Visible evidence is preserved
in `docs/screenshots/clippy-global-controller.jpg`.

The phase 3 MCP implementation is recorded in `docs/specs/mcp/README.md`.
Its initialize path supports both MCP `2025-06-18` (used by the current
Codex desktop app-server) and `2025-11-25`, and accepts standard request
`_meta` fields on lifecycle and tool requests. Tool arguments remain
strictly allowlisted.
Phase 4 remains tightly scoped read-only XP automation with explicit
confirmation; Phase 5 is separate Word-native Assistant integration with
intentional visible handoff; and Phase 6 is final XP validation and
documentation synchronization. Agent request completion/error observation
must be designed before actions are represented as completed rather than
queued.

The persistent phase 3 service's disconnect lifecycle was hardened on
2026-09-08. TCP client cleanup now precedes any final Agent teardown, the
visible COM controller persists across successive MCP protocol sessions, and
the stdio bridge has a bounded close. Host and XP fake-COM tests plus repeated
live bridge sessions verified that closed clients leave only normal
`TIME_WAIT` entries and the same XP listener remains ready for the next
connection.

Metadata-only MCP diagnostics were added on 2026-09-08. Host and XP Python
3.4 tests verify that method/tool routing and transport progress are logged,
while tool arguments and message text are absent. A live standalone bridge
session verified initialize, `tools/list`, `clippy.animations`, both JSONL log
paths, and clean return to the listener's next `accept()`. A second session
left the connection idle for more than ten seconds and exposed a transport
defect: the bridge's socket retained its connect timeout, so the receive thread
exited on idle and could not forward a later response. The bridge now clears
that timeout after connecting. The short connection deadline and bounded EOF
shutdown remain unchanged, while established sessions can stay idle until
Codex closes stdin or the XP service closes the socket.

The same persistent service now pumps Agent COM while its MCP sockets are idle.
Two fake-socket regressions cover an open client that pauses before a later
request and a disconnected client followed by an idle listener. Live XP
acceptance queued `show`, `move`, `Greeting`, and `think`, then sent no MCP
requests for twelve seconds: Clippy advanced from the greeting into the Think
animation, and a later `tools/list` response returned normally. Clippy also
continued animating after that client disconnected and the listener returned
to `accept()`.

Historical Word integration track

1. Native Clippy control

Prove native C++ can programmatically drive the already-working Microsoft Agent installation:

Show
Speak
Play animation
Move
Ask questions

2. Identify Office’s input control — resolved

A before/after WindowWatch comparison identified the complete window subtree
that appears when Word opens the original “What would you like to do?”
balloon:

```text
MSOBALLOON                         agentsvr.exe
└─ MsoBalloonChild                 WINWORD.EXE
   ├─ RichEdit20W   Control ID 6   query editor
   ├─ Button        Control ID 7   owner-drawn command button
   └─ Button        Control ID 8   owner-drawn command button
```

The authentic input control is the `RichEdit20W` child with control ID `6`.
When the balloon first opens, it has keyboard focus and contains the prompt
`Type your question here and then click Search.` The “What would you like to
do?” heading is painted by the balloon implementation rather than exposed as a
separate `Static` control.

The process boundary is significant. The top-level `MSOBALLOON` shell belongs
to `agentsvr.exe`, while `MsoBalloonChild`, the query editor, and both buttons
belong to `WINWORD.EXE` and run on Word’s UI thread. The separate
`AgentAnimBalloon` window is used for Microsoft Agent speech/thought output and
is not this Office help UI.

Captured HWNDs and process IDs are transient and must never be hard-coded. The
shim should locate the editor from its class hierarchy and stable control ID,
then confirm that the content window belongs to the active Word process:

```cpp
HWND balloon = FindWindowW(L"MSOBALLOON", NULL);
HWND content = FindWindowExW(
    balloon, NULL, L"MsoBalloonChild", NULL);
HWND query = GetDlgItem(content, 6);
```

An external utility can inspect or replace the editor text with normal window
messages. Capturing and overriding the built-in submission behavior should be
implemented in `ClippyShim.dll`, loaded inside `WINWORD.EXE`, so it can safely
subclass `MsoBalloonChild` and observe its message flow. Attempting to replace
the window procedure directly from an external process is not a safe design.

The two command buttons use `BS_OWNERDRAW`, so their text is not available
through `WM_GETTEXT`. Visible UI verification mapped control ID `7` to Options
and control ID `8` to Search. The Enter and Search interception paths should
therefore consume only control `8`; control `7` must continue to Office's
original window procedure.

3. Capture a real Office query

The critical reverse-engineering milestone:

User types into authentic Clippy UI:
"why is my build failing?"
             ↓
ClippyShim logs:
QUERY=why is my build failing?

At this point, the hardest historical integration problem is essentially solved.

`ClippyShim.dll` now implements this path as a native x86 Word COM add-in. It
registers per-user as `ClippyShim.Connect`, receives Word's
`_IDTExtensibility2::OnConnection`, and installs a UI-thread timer without
creating a window of its own. The timer locates the live Assistant hierarchy
by class name, control ID, and current process ownership, then subclasses both
`MsoBalloonChild` and its `RichEdit20W` query editor in-process.

The interception path handles an unmodified Enter key in the editor and the
`BN_CLICKED` notification from control ID `8`, the right-hand Search button.
It rejects the original placeholder and empty input, logs the captured text to
`C:\clippy\ClippyShim.log`, suppresses the legacy Help submission, and changes
the editor immediately to `Asking Clippy...`. Milestone 4 owns the network
request and renders the returned text under the user-facing `Clippy says:`
heading with Word's native `Assistant.NewBalloon` API. Transport terms such as
"host" stay out of visible Assistant copy.

The add-in never hard-codes window handles or process IDs, restores each
original window procedure before detaching, and keeps all Word/Assistant COM
calls on Word's UI thread. The original local-echo acceptance test is
preserved in `docs/screenshots/clippy-query.png` and
`docs/screenshots/clippy-echo.png`; the current host-bridge acceptance evidence
is described under milestone 4.

4. Bridge Clippy to macOS — AI rich-text path implemented

The captured query now crosses the VM boundary as an HTTP request:

XP → Mac
{"text":"why is my build failing?"}

and receives a model-generated reply:

Mac → XP
{
  "text": "Your linker cannot find the library named in the build settings."
}

`ClippyShim.dll` starts a short-lived background worker for each accepted
query, UTF-8 encodes it, and posts it to
`http://10.0.2.2:3210/message`. `10.0.2.2` is the host gateway exposed to this
XP guest by its current QEMU/UTM network. Connect and send operations retain
three-second timeouts; receive operations allow 60 seconds for model
generation. Responses remain limited to 64 KiB. The worker never calls Word
COM; it places the response into synchronized state, and the existing Word
UI-thread timer formats and displays it through `Assistant.NewBalloon`. This
preserves the apartment boundary and avoids freezing Word during network I/O.

The macOS side lives under `server/` and is a strict TypeScript Node HTTP
server. It binds `0.0.0.0:3210`, accepts only `POST /message`, validates a
non-empty string `text` property, and sends it to OpenAI's Responses API. The
official JavaScript SDK reads `OPENAI_API_KEY`; `CLIPPY_OPENAI_MODEL` selects
the model and defaults to `gpt-5.6-luna`. Requests use low reasoning effort,
low text verbosity, a 512-token output ceiling, no tools, no previous response,
and no response storage. Code-managed instructions keep replies short,
accurate, lightly in character, and within the supported Markdown-lite subset.
Its bind address and port remain configurable with `CLIPPY_SERVER_BIND` and
`CLIPPY_SERVER_PORT`; the XP endpoint remains fixed.

Successful responses and failures both return to the UI thread and are
rendered in native Office Assistant balloons. Malformed XP requests return
HTTP 400, model timeouts return 504, and other model failures return 502. The
shim parses safe JSON error text and presents it under `Clippy couldn't
answer:`. The response parser handles UTF-8 and JSON string escapes.

For pending-request feedback, the shim starts Office Assistant's authentic,
looping `msoAnimationThinking` animation immediately after accepting a query.
The call is late-bound on Word's UI thread, just like the response balloon; the
network worker never touches Office COM. When either a response or an error
reaches the UI timer, the shim resets the Assistant to `msoAnimationIdle`
before rendering the result. Animation automation is intentionally
best-effort: a character or Office automation failure is logged but does not
cancel an otherwise valid host request.

Rich response rendering is an XP compatibility responsibility. The HTTP
contract remains `{"text":"..."}`, and the server returns Markdown-lite rather
than Office-specific control codes. On Word's UI thread, `ClippyShim.dll` maps:

* `**strong**` and level 1-3 headings to dark-blue underlined text
* `*emphasis*` to underlined text
* inline code to dark-blue text with visible backticks, and fenced code to
  dark-blue lines with a visible `|` gutter
* links to a blue underlined label followed by the visible URL
* one final list of at most five items to native Office bullet or number labels

Office Assistant balloons do not expose arbitrary HTML, RTF, fonts, or inline
bold. The mapping uses only the documented native `{ul}` and `{cf}` directives
plus `BalloonType` and `Labels`. Model-provided braces are changed to full-width
braces before rendering, preventing an answer from injecting Office directives
or local BMP/WMF references. Unsupported or malformed Markdown remains readable
literal text. Native lists are limited to the five labels supported by Office;
a longer, nested, mixed, or non-final list remains inline text.

The server is stateless at this checkpoint. Conversation memory, tools,
host-selected animation/action data, and action buttons remain milestone 5/6
work. The pending Thinking animation is a fixed XP-side interaction state, not
a new host protocol field.

The original echo-path screenshots remain under `docs/screenshots/`. The live
AI path was freshly verified on the visible UTM desktop with the query `In five
words, what is retro computing?` and the native response `Old computers,
software, and games nostalgia.` Initial rich Markdown rendering is captured in
`docs/screenshots/clippy-rich-text.png`; the revised code styling is captured in
`docs/screenshots/clippy-rich-code.png`; and the current `Clippy says:` heading
is captured in `docs/screenshots/clippy-says.png`.

5. Give Clippy tools

The modern agent gains controlled access to:

* coding environment
* shell
* files
* browser
* APIs
* email/calendar
* other applications

Clippy becomes the interface through which these capabilities are invoked.

6. General-purpose AI Clippy

The final experience:

📎 What would you like to do?
> Figure out why this project won't compile.
📎 I found the problem. You're linking the wrong library.
   [Fix it]   [Explain]   [Ignore]

The pixels, animations, Office UI, and character runtime are authentic.

The intelligence behind them is from 2026.

⸻

Design Philosophy

The project deliberately avoids making a Clippy-themed AI assistant.

The goal is to take the actual software architecture Microsoft shipped decades ago and replace the obsolete intelligence behind it while preserving as much of the original interface and behavior as possible.

In other words:

Don’t recreate Clippy. Possess him.
