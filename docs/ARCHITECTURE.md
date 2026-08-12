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

Current Status — 2026-08-11

The first four milestones and the first AI response checkpoint are complete:

* native code can control the installed Clippit character
* the authentic Office Assistant query editor and Search command are identified
* `ClippyShim.dll` intercepts Search and unmodified Enter inside Word
* XP posts the captured text to the macOS TypeScript server
* the server sends the text to OpenAI's Responses API with `gpt-5.6-luna`
* Word renders the model-generated text in a native Assistant balloon

The complete live model flow has been visibly verified on the XP desktop and
independently verified at both HTTP boundaries. This checkpoint is deliberately
single-turn and text-only. Milestones 5 and 6 remain open: the next work is to
add controlled host tools, then conversation/session state and richer Clippy
animation/action responses. Making the XP endpoint configurable also remains a
protocol follow-up.

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
│  Office XP                                                   │
│      │                                                       │
│      ▼                                                       │
│  Original Office Assistant UI                                │
│      │                                                       │
│      ▼                                                       │
│     📎 Clippy                                                │
│      │                                                       │
│      ├── ClippyShim.dll                                      │
│      │      ├── capture user input                           │
│      │      ├── send requests to Mac                         │
│      │      └── render responses/actions                     │
│      │                                                       │
│      └── Microsoft Agent / Office Assistant APIs             │
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

Microsoft Agent itself supplies the animated character, while Office adds richer interactive UI on top.

ClippyShim.dll

A native Win32/x86 component loaded into Office is the primary XP integration
layer.

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
the editor immediately to `Asking the Clippy host...`. Milestone 4 owns the
network request and renders the returned text with Word's native
`Assistant.NewBalloon` API.

The add-in never hard-codes window handles or process IDs, restores each
original window procedure before detaching, and keeps all Word/Assistant COM
calls on Word's UI thread. The original local-echo acceptance test is
preserved in `docs/screenshots/clippy-query.png` and
`docs/screenshots/clippy-echo.png`; the current host-bridge acceptance evidence
is described under milestone 4.

4. Bridge Clippy to macOS — AI text path implemented

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
UI-thread timer displays it through `Assistant.NewBalloon`. This preserves the
apartment boundary and avoids freezing Word during network I/O.

The macOS side lives under `server/` and is a strict TypeScript Node HTTP
server. It binds `0.0.0.0:3210`, accepts only `POST /message`, validates a
non-empty string `text` property, and sends it to OpenAI's Responses API. The
official JavaScript SDK reads `OPENAI_API_KEY`; `CLIPPY_OPENAI_MODEL` selects
the model and defaults to `gpt-5.6-luna`. Requests use low reasoning effort,
low text verbosity, a 512-token output ceiling, no tools, no previous response,
and no response storage. Code-managed instructions keep replies plain-text,
short, accurate, and lightly in character for the Office balloon. Its bind
address and port remain configurable with `CLIPPY_SERVER_BIND` and
`CLIPPY_SERVER_PORT`; the XP endpoint remains fixed.

Successful responses and failures both return to the UI thread and are
rendered in native Office Assistant balloons. Malformed XP requests return
HTTP 400, model timeouts return 504, and other model failures return 502. The
shim parses safe JSON error text and presents it under `Clippy couldn't
answer:`. The response parser handles UTF-8 and JSON string escapes.

The server is stateless and text-only at this checkpoint. Conversation memory,
tools, optional animation data, and action buttons remain milestone 5/6 work.
The original echo-path screenshots remain under `docs/screenshots/`; the live
AI path was freshly verified on the visible UTM desktop with the query `In five
words, what is retro computing?` and the native response `Old computers,
software, and games nostalgia.`

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
