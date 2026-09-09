# Clippy Possession development log

## 2026-08-11 - user-facing balloon headings

Removed transport terminology from visible Assistant copy. Successful response
balloons now use `Clippy says:` instead of `Clippy host replied:`, and the query
editor shows `Asking Clippy...` while a request is pending instead of `Asking
the Clippy host...`. Internal log and server terminology remain unchanged.

Mirrored `src/addin/ClippyShim.cpp` to XP, ran the complete Visual C++ 2010
`build.bat`, and registered the replacement DLL. Visible acceptance launched
Word through Windows Key+R, opened the question balloon with F1, and submitted
`Reply with exactly: Header wording verified.` using unmodified Enter. The shim
log recorded `QUERY`, the Thinking animation, and `RESPONSE text=Header wording
verified.` A fresh screenshot visibly confirmed the `Clippy says:` heading with
the same response at `docs/screenshots/clippy-says.png`. UTM Capture Input was
not used. Updated `AGENTS.md` so future acceptance checks require the new
heading.

## 2026-08-11 - revised code-span styling

Changed the Markdown code mapping after visual feedback showed that Office's
dark-cyan palette entry reads green against the yellow Assistant balloon and,
after the Markdown delimiters were removed, did not clearly look like code.
Inline code now renders in dark blue with its backticks visible. Fenced code
uses the same dark blue with a visible `|` gutter on each content line. This is
the clearest available approximation because `Assistant.NewBalloon` cannot set
a monospace font or code-span background.

Mirrored `src/addin/ClippyShim.cpp` to XP, ran the complete `build.bat`, and
registered the replacement DLL. Visual acceptance launched Word from Windows
Key+R, opened the authentic question balloon with F1, and submitted:

```text
Reply exactly with this Markdown: **Run checks first.** Then on a new
paragraph: Use `npm test`, then `npm run build`.
```

The live OpenAI response preserved that Markdown. A fresh screenshot showed
both commands in dark blue with visible backticks, distinct from the
blue-underlined strong sentence above them. The shim log recorded the query,
Thinking animation, and exact response, and the final screenshot is stored at
`docs/screenshots/clippy-rich-code.png`. UTM Capture Input was not used.

## 2026-08-11 - native rich-text OpenAI replies

Added a safe Markdown-lite renderer to `ClippyShim.dll`. The host still returns
the stable `{"text":"..."}` response, while the XP compatibility layer maps
strong emphasis and headings to dark-blue underlined text, emphasis to
underlining, inline/fenced code to dark cyan, and links to an underlined blue
label followed by the visible URL. One final Markdown list of at most five
items becomes Office's native `BalloonType` bullet or number labels.

The renderer emits only Office's documented `{ul}` and `{cf}` directives.
Every brace originating in the model response becomes a full-width brace before
rendering, so model output cannot inject Office formatting or the supported
local `{bmp ...}` / `{wmf ...}` picture references. Longer, nested, mixed, or
non-final lists stay readable inline rather than being moved into native labels.
The host instructions now invite only the supported Markdown subset and
explicitly forbid HTML, images, nested lists, tables, and raw Office
directives. Added `scripts/diagnostics/test_balloon_markup.vbs` as a direct
visual probe of the installed Office formatting behavior.

Verified the host and complete XP build:

```sh
cd server
npm test
npm run build

scp src/addin/ClippyShim.cpp windows-xp:'C:/clippy/src/addin/ClippyShim.cpp'
ssh windows-xp 'cd /d C:\clippy && build.bat'
ssh windows-xp 'regsvr32 /s C:\clippy\build\ClippyShim.dll'
```

All seven TypeScript tests passed, and Visual C++ 2010 produced the three
expected x86/XP targets without warnings. Before integrating the parser, the
new diagnostic visibly confirmed on Office XP that `{ul}`, `{cf}`, and native
bullet labels render in an authentic Assistant balloon.

The final visible acceptance used Windows Key+R, `winword.exe`, F1, and
unmodified Enter. The submitted request asked OpenAI for this exact Markdown:

```text
**Rich text works.**

- `inline code`
- *underlined emphasis*
```

The live host returned that text, `C:\clippy\ClippyShim.log` recorded the query,
Thinking animation, and response, and a fresh UTM screenshot showed `Rich text
works.` in blue underline, `inline code` in dark cyan, and `underlined
emphasis` underlined as two native bullet labels. The visible result is stored
at `docs/screenshots/clippy-rich-text.png`. UTM Capture Input was not used.

Known limitations: Office Assistant has no arbitrary HTML/RTF, inline bold,
font, or clickable-link support. Rich formatting is intentionally limited to
the native Office features above, list labels are capped at five, and replies
remain single-turn with no tools, streaming, or structured animation/actions.

## 2026-08-11 - Thinking animation during host requests

Added pending-request feedback to `ClippyShim.dll` using Office Assistant's
authentic, looping `msoAnimationThinking` animation. Query submission still
starts the HTTP exchange on a background worker, while the Word UI thread now
sets `Assistant.Visible` and `Assistant.Animation` immediately after accepting
the query. The existing UI timer resets the animation to `msoAnimationIdle`
before showing either a successful response or an error balloon. Animation
failures are logged and remain non-fatal to the host request.

Verified the complete XP build after closing the known visible demo Word
instance and copying the changed source:

```sh
scp src/addin/ClippyShim.cpp windows-xp:'C:/clippy/src/addin/ClippyShim.cpp'
ssh windows-xp 'cd /d C:\clippy && build.bat'
ssh windows-xp 'regsvr32 /s C:\clippy\build\ClippyShim.dll'
```

The VS2010 x86/XP build produced all three expected targets and reported
`OK: built C:\clippy\build\ClippyShim.dll`. The replacement add-in registered
successfully. A macOS `curl` request to the live TypeScript server also returned
HTTP 200 before the visible test.

For visible acceptance, launched `winword.exe` through Windows Key+R on the UTM
desktop, opened the authentic Assistant query balloon with F1, entered `In one
short sentence, why does Clippy think before answering?`, and submitted with
unmodified Enter. A fresh screenshot during the request showed Clippy in the
Thinking animation with the editor text `Asking the Clippy host...`; a later
screenshot showed the native `Clippy host replied:` balloon with `Clippy thinks
before answering to make sure the help is accurate and useful.` The shim log
correlated the interaction at 22:19:56-22:19:58:

```text
QUERY source=Enter key text=In one short sentence, why does Clippy think before answering?
ANIMATION started Thinking
ANIMATION stopped Thinking
RESPONSE text=Clippy thinks before answering to make sure the help is accurate and useful.
```

The animation remains a fixed XP-side pending state; the host response protocol
still carries text only. Reset occurs on the 200 ms Word UI poll, and a failure
to automate the selected Office Assistant is logged rather than failing the
network request.

## 2026-08-11 - toolchain inventory

- Host: Microsoft Windows XP 5.1.2600, x86.
- Working directory: `C:\clippy` (initially empty).
- Visual C++: Visual Studio 2010, with `cl.exe` at
  `C:\Program Files\Microsoft Visual Studio 10.0\VC\bin\cl.exe`.
- Environment setup: `C:\Program Files\Microsoft Visual Studio 10.0\VC\vcvarsall.bat x86`.
- Windows SDKs installed: v7.0A and v7.1. Visual Studio reports v7.0A as its
  current SDK (`7.0.30319`).
- Microsoft Agent control: `Agent.Control.2`, CLSID
  `{D45FD31B-5C6E-11D1-9EC1-00C04FD7081F}`, implemented by
  `C:\WINDOWS\msagent\agentctl.dll` with Apartment threading.
- Clippit character file:
  `C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS`.

Discovery commands:

```bat
ver
set
dir C:\clippy
reg query "HKLM\SOFTWARE\Microsoft\VisualStudio\10.0\Setup\VC" /v ProductDir
reg query "HKLM\SOFTWARE\Microsoft\Microsoft SDKs\Windows" /s
reg query "HKCR\Agent.Control.2" /s
reg query "HKCR\CLSID\{D45FD31B-5C6E-11D1-9EC1-00C04FD7081F}" /s
dir "C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS"
```

## Implementation

`src/tools/ClippyProbe.cpp` uses native COM (`CoInitializeEx`, `CLSIDFromProgID`,
`CoCreateInstance`, and `IDispatch`). Late binding keeps the probe independent
of generated ActiveX wrappers and does not alter Microsoft Agent registration.

The probe connects the control, loads the installed Clippit character, and
invokes `Show` and `Think`. It prints a checkpoint after each successful COM
operation and returns a distinct nonzero exit code on failure.

Build command:

```bat
cd /d C:\clippy
build.bat
```

Run command:

```bat
C:\clippy\build\ClippyProbe.exe
```

## Build and invocation result

The build script was invoked by absolute path from the SSH session, confirming
that it reliably changes to its own directory before compiling:

```text
src\tools\ClippyProbe.cpp
OK: built C:\clippy\build\ClippyProbe.exe
```

`dumpbin /headers` confirms the produced binary is PE32, machine `14C (x86)`,
with operating-system and subsystem versions both set to `5.01` for Windows XP.
The Visual C++ tools report version `10.00.30319.01`.

Native invocation completed with exit code 0:

```text
OK: created Agent.Control.2.
OK: connected to Microsoft Agent.
OK: loaded C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS.
OK: invoked Clippy.Show().
OK: invoked Clippy.Think().
```

No Microsoft Agent registration was changed. No SSH desktop-isolation work was
performed.

## 2026-08-11 - WindowWatch

Added `src/tools/WindowWatch.cpp`, a windowless Win32 GUI-subsystem process that captures
the complete desktop window tree when F11 is pressed. It creates new UTF-8
`C:\clippy\dumps\dump-NNN.txt` files with restart-safe numbering and appends a
summary row to `C:\clippy\dumps\index.txt`.

Each record includes HWND, PID, TID, parent, owner, class, text, control ID,
visibility, enabled state, rectangle, style, and extended style. Child HWNDs
are recursively indented beneath their parents. Foreground focus is read with
`GetGUIThreadInfo`; the utility never creates a window or calls a foreground,
activation, or focus-changing API.

WindowWatch registers bare F11 with `RegisterHotKey`. F11 replaced the original
F12 choice because Windows XP reserves F12 for debugger use. No keyboard hook
or foreground-changing API is needed.

The VS2010 x86 `/MT` build completes cleanly. `dumpbin /headers` reports PE32,
machine `14C (x86)`, Windows GUI subsystem, and OS/subsystem version 5.01.

Interactive validation launched WindowWatch through XP's Start > Run while
Word was open. Two hotkey presses three seconds apart produced sequential dumps
with the same Word foreground HWND:

```text
003 | ... | 0x000803EA | OpusApp | Document1 - Microsoft Word
004 | ... | 0x000803EA | OpusApp | Document1 - Microsoft Word
```

After terminating WindowWatch, rebuilding, and launching it again, the next
two presses continued with `dump-005.txt` and `dump-006.txt`; both again
recorded the same Word foreground HWND. A field-count audit of `dump-005.txt`
found 522 HWND records and exactly 522 occurrences of every required field.

## 2026-08-11 - built-in Office Assistant help balloon discovery

Compared `docs/captures/windowwatch/dump-001.txt`, captured while the Word
document was active, with `docs/captures/windowwatch/dump-002.txt`, captured
after opening Clippy's built-in "What would you like to do?" help balloon. The
second capture changed the foreground window to
`MSOBALLOON` and placed keyboard focus in its query editor. The complete new
visible subtree was:

```text
MSOBALLOON                         agentsvr.exe
└─ MsoBalloonChild                 WINWORD.EXE
   ├─ RichEdit20W   Control ID 6   query editor
   ├─ Button        Control ID 7   owner-drawn command button
   └─ Button        Control ID 8   owner-drawn command button
```

The capture-specific handles and ownership were:

```text
0x000303A2  MSOBALLOON       PID 2476 (agentsvr.exe), foreground
0x00050426  MsoBalloonChild  PID 2364 (WINWORD.EXE)
0x00020450  RichEdit20W      PID 2364, Control ID 6, focused
0x0002045E  Button           PID 2364, Control ID 7
0x0002045C  Button           PID 2364, Control ID 8
```

`tasklist` independently confirmed PID 2476 as `agentsvr.exe` and PID 2364 as
`WINWORD.EXE`. The editor's window text was `Type your question here and then
click Search.` The two buttons have style `BS_OWNERDRAW` (`0x0000000B` in the
button-specific style bits), which explains why their `WM_GETTEXT` values are
empty. Their semantic mapping still needs to be measured rather than inferred
from screen position.

The heading "What would you like to do?" is rendered by the balloon rather
than exposed as a separate `Static` child window. The similarly named
`AgentAnimBalloon` belongs to Microsoft Agent's speech/thought UI and is not
the Office help dialog.

The HWND and PID values above are observations from one run and must not be
hard-coded. A runtime locator should use the class hierarchy and stable control
ID, then verify that the content window belongs to the active Word process:

```cpp
HWND balloon = FindWindowW(L"MSOBALLOON", NULL);
HWND content = FindWindowExW(
    balloon, NULL, L"MsoBalloonChild", NULL);
HWND query = GetDlgItem(content, 6);
```

This ownership split determines the likely implementation boundary. The shell
window is hosted by `agentsvr.exe`, but `MsoBalloonChild`, the query editor,
and both command buttons run on Word's UI thread. An external diagnostic can
read or set the editor with normal window messages. Intercepting the built-in
submit path should instead be done by code loaded into `WINWORD.EXE`, such as a
Word COM add-in that subclasses `MsoBalloonChild`, because replacing a window
procedure across process boundaries is not a safe external-helper operation.

Office also exposes `Application.Assistant.NewBalloon` as the supported path
for presenting a custom Assistant balloon. That is suitable for replacing the
experience under our own command, but it does not by itself override the
built-in `Assistant.Help` action that creates this search balloon.

Next diagnostic step: build an in-process message probe for
`MsoBalloonChild` that records the `WM_COMMAND` and edit-notification sequence
for Enter and for each owner-drawn button. This will establish the exact roles
of control IDs 7 and 8 and the narrowest interception point before any behavior
is changed.

## 2026-08-11 - ClippyShim Word COM add-in

Added `src/addin/ClippyShim.cpp`, a dependency-light native COM server implementing
Office's `_IDTExtensibility2` ABI. `DllRegisterServer` writes the COM class and
`ClippyShim.Connect` ProgID under the current user's merged `HKCR` view, plus
Word's per-user COM add-in registration with `LoadBehavior=3`. No administrator
registration or Office installation changes are required.

On connection, the DLL creates a windowless UI-thread timer. It discovers the
visible `MSOBALLOON` / `MsoBalloonChild` / `RichEdit20W` hierarchy, verifies
that the content belongs to the current `WINWORD.EXE`, and subclasses the
content and editor controls in-process. Enter and button ID `8` capture a
non-placeholder query, append a UTF-8 `QUERY` record to
`C:\clippy\ClippyShim.log`, and suppress the old Office Help submission. The
captured text is echoed with `Application.Assistant.NewBalloon`; immediate
`You said: ...` editor text remains as a visible fallback.

The VS2010 x86 `/MT /LD` build succeeded after adding `advapi32.lib` for the
per-user registry functions. The generated DLL exports the standard COM entry
points through `src/addin/ClippyShim.def`. Silent `regsvr32` registration
succeeded, and
the resulting Word key was verified as:

```text
HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\ClippyShim.Connect
    FriendlyName     REG_SZ     Clippy Possession Shim
    Description      REG_SZ     Intercepts Office Assistant questions for Clippy Possession.
    LoadBehavior     REG_DWORD  0x3
    CommandLineSafe  REG_DWORD  0x0
```

An SSH-side `GetObject(, "Word.Application")` test returned `MK_E_UNAVAILABLE`
(`0x800401E3`), confirming that the existing Word instance is isolated on the
visible UTM desktop and cannot be hot-loaded from the SSH desktop. Interactive
acceptance therefore still requires a fresh Word launch through UTM followed
by screenshots of the authentic query and echo balloons.

Visible UTM inspection subsequently resolved the owner-drawn button mapping:
control ID `7` is Options and control ID `8` is Search. The shim was corrected
to intercept only ID `8`, leaving Options on Office's original behavior.

After rebuilding, Word was closed and relaunched from the visible XP desktop.
The add-in loaded at startup, attached both in-process hooks when F1 opened the
built-in Assistant prompt, intercepted the Search button, and displayed the
echo with Word's native Assistant balloon. The completed log sequence was:

```text
ADDIN connected
HOOK attached to Office Assistant query balloon
QUERY source=Search button text=echo from ClippyShim
ECHO shown with Assistant.NewBalloon
```

Fresh screenshots preserve the authentic input and output states at
`docs/screenshots/clippy-query.png` and `docs/screenshots/clippy-echo.png`.

## 2026-08-11 - repository organization

Grouped the native implementation under `src/addin`, native probes under
`src/tools`, VBScript utilities under `scripts/demos` and
`scripts/diagnostics`, and curated WindowWatch evidence under
`docs/captures/windowwatch`. Generated binaries remain isolated in the ignored
`build` directory.

The root `build.bat` now compiles from the organized source paths while keeping
the existing output names and `C:\clippy\build` deployment contract. The same
tree was mirrored to XP and the complete VS2010 x86 build passed.

Updated `AGENTS.md` to require consulting `docs/ARCHITECTURE.md`, recording
meaningful work in this development log, validating the relevant XP path, and
committing each coherent source-and-documentation change without generated or
unrelated files.

## 2026-08-11 - XP-to-macOS TypeScript echo bridge

Added `server/src/server.ts`, a strict TypeScript Node HTTP server that listens
on `0.0.0.0:3210`. It accepts only `POST /message`, requires a non-empty JSON
`text` string, prints the received message, and returns the same text as JSON.
The npm project has no runtime dependencies; TypeScript and Node types are
development-only dependencies used to build `server/dist/server.js`.

Replaced ClippyShim's local echo with a Winsock HTTP client. Captured questions
are UTF-8/JSON encoded and sent to `10.0.2.2:3210`, the host gateway observed
from this XP VM. Network I/O runs on a short-lived worker thread with
three-second timeouts and a 64 KiB response cap. The worker publishes its
result through a critical section, while Word's existing UI-thread timer alone
continues to make Office COM calls and show the returned text with
`Assistant.NewBalloon`. Connection and protocol errors are also shown in a
native balloon. Added `scripts/diagnostics/test_host_bridge.vbs` for an
independent XP-side transport check.

Host build and loopback validation:

```text
cd server
npm install
npm run build
npm start
curl -H 'Content-Type: application/json' \
  --data '{"text":"hello from host test"}' \
  http://127.0.0.1:3210/message
=> {"text":"hello from host test"}
```

The independent XP diagnostic was copied to Brett's resolved `%TEMP%` and run
with `cscript.exe //nologo %TEMP%\test_host_bridge.vbs`. It returned:

```text
HTTP 200
{"text":"hello from Windows XP"}
```

After mirroring `ClippyShim.cpp` and `build.bat` to `C:\clippy`, the complete
VS2010 x86 `/W4` build passed without warnings. The rebuilt DLL was registered
per-user, and Word was launched from the visible XP desktop. F1 opened the
authentic Assistant query balloon; submitting `echo through TypeScript` with
Search produced this correlated path:

```text
host: Clippy: echo through TypeScript
XP:   QUERY source=Search button text=echo through TypeScript
XP:   RESPONSE text=echo through TypeScript
XP:   RESPONSE shown with Assistant.NewBalloon
```

Fresh visible evidence is preserved in
`docs/screenshots/clippy-bridge-query.png` and
`docs/screenshots/clippy-bridge-response.png`. The latter shows the native
balloon heading `Clippy host replied:` and the echoed text. UTM's Capture Input
control was not used.

Current limitation: this milestone has a fixed XP endpoint and a text-only
response contract. The TypeScript server has bind/port environment settings,
but configuring the XP endpoint and consuming an optional animation command
remain future work.

## 2026-08-11 - current workflow and milestone status

Consolidated the verified build, deployment, run, transport-diagnostic, and
visible acceptance sequence in `AGENTS.md`. Future work now has one canonical
workflow covering the macOS TypeScript build and server, native source mirroring
to `C:\clippy`, the VS2010 x86 build, per-user add-in registration, the
independent XP HTTP diagnostic, and the UTM-only Word/Clippy acceptance path.

Updated `docs/ARCHITECTURE.md` with the current project checkpoint. Milestones
1 through 4 are complete: native Clippit control, authentic Office query UI
discovery, in-process query interception, and the XP-to-macOS echo bridge have
all been implemented and visibly verified. The TypeScript server remains an
echo handler. Milestones 5 and 6—host tools and the general-purpose agent—are
not yet implemented. The next protocol work is configurable XP endpoint data
and optional animation/action fields in host responses.

Re-ran the commands documented in the workflow. `npm run build` completed the
strict TypeScript build, the host curl request returned
`{"text":"hello from macOS"}`, and `ssh windows-xp 'cd /d C:\clippy &&
build.bat'` rebuilt all three XP binaries without warnings. The independent
`test_host_bridge.vbs` diagnostic then returned HTTP 200 and
`{"text":"hello from Windows XP"}` from the running host server. No generated
binaries, server output, or runtime logs are part of this documentation change.

## 2026-08-11 - OpenAI-powered Clippy replies

Replaced the host echo handler with a single-turn OpenAI Responses API call
through the official JavaScript SDK. The default model is `gpt-5.6-luna`, with
an environment override through `CLIPPY_OPENAI_MODEL`; startup now fails before
binding when `OPENAI_API_KEY` is absent. Requests use low reasoning effort,
low verbosity, a 512-token output ceiling, no tools or prior response, and
`store: false`. Code-managed instructions constrain replies to concise,
plain-text, lightly playful Clippy answers suitable for an Office balloon.

The HTTP success contract remains `200 {"text":"..."}`. The reply generator is
injectable for network-free tests. Bad input returns 400, OpenAI timeouts return
504, and other upstream or empty-output failures return 502 with safe error
text. The XP shim now keeps three-second connect/send limits but allows a
60-second receive, parses JSON `error` strings on non-200 responses, and shows
them under `Clippy couldn't answer:`. The XP VBScript diagnostic uses the same
60-second receive allowance. Repository `.env` files are ignored so API keys
cannot be added accidentally.

Host validation:

```text
cd server && npm test
=> 7 tests passed; 0 failed

env -u OPENAI_API_KEY npm start
=> exit 1: OPENAI_API_KEY must be set before starting the Clippy host server

curl --fail-with-body -H 'Content-Type: application/json' \
  --data '{"text":"In one short sentence, why was the original Clippy infamous?"}' \
  http://127.0.0.1:3210/message
=> {"text":"Clippy was infamous for interrupting users with unsolicited, often unhelpful advice."}
```

A controlled invalid credential exercised the real SDK error path: host curl
returned HTTP 502 with `{"error":"OpenAI request failed"}`, and the XP
diagnostic received the same response. A deterministic injected generator then
returned HTTP 200 through the XP diagnostic, proving that tests and acceptance
do not depend on a live model for protocol coverage.

Mirrored `src/addin/ClippyShim.cpp` and the updated diagnostic to XP. The full
`C:\clippy\build.bat` VS2010 x86 build produced all three binaries without
warnings, and silent per-user registration preserved `LoadBehavior=3`. With a
live API key, `cscript.exe //nologo %TEMP%\test_host_bridge.vbs` returned:

```text
HTTP 200
{"text":"Hello! It's great to hear from Windows XP. How can I help you today?"}
```

The visible UTM acceptance used Windows Key+R to launch Word, F1 to open the
authentic Assistant question editor, and unmodified Enter to submit `In five
words, what is retro computing?`. The live `gpt-5.6-luna` reply appeared in a
fresh native balloon under `Clippy host replied:` as `Old computers, software,
and games nostalgia.` The XP log correlated `QUERY` and `RESPONSE`; a separate
controlled failure visibly produced `Clippy couldn't answer: OpenAI request
failed`. UTM's Capture Input control was not used.

Known limitations: replies are single-turn and text-only; there is no memory,
tool execution, animation/action data, streaming, or configurable XP endpoint.
The 150-word instruction is a model constraint rather than a server-side word
truncation, while the 512-token API ceiling remains the hard output bound.

## 2026-08-19 - persistent desktop-wide Microsoft Agent controller foundation

Added `xp/clippy_agent.py`, a Python 3.4-compatible controller that holds one
real `Agent.Control.2` connection and the installed Office XP Clippit character
for the lifetime of a stdin session. This is the new primary desktop-wide
character boundary; the existing `ClippyShim.dll` remains the working
Word-specific integration. The controller supports show, hide, move, speak,
think, animation enumeration, and guarded play. Play accepts only exact names
enumerated from `C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS` at
startup.

Added a narrow newline JSON-RPC 2.0 adapter with no arbitrary shell,
filesystem, input injection, generic COM, or broad desktop-control method.
Requests reject extra parameters, bound text and coordinate inputs, return
standard parse/request/method/parameter errors, and keep internal COM error
details on stderr. Agent actions return `queued` rather than claiming
completion because Microsoft Agent returns asynchronous request objects whose
final status is not yet observed. Added a reusable visible smoke script and a
two-request redirected protocol probe.

The first attempted dependency, `comtypes==1.2.1`, is the last release that
supports Python 3.4 and imported successfully on XP. Both dynamic and generated
dispatch created `Agent.Control.2`, but setting `Connected = True` hung. The
validated path uses the official archived 32-bit pywin32 build 220 installer
for Python 3.4. The legacy installer had to be completed on the visible XP
desktop; it did not honor quiet mode over SSH. Agent clients launched through
SSH also hung during `Connected`, so GUI character hosts must follow the normal
visible-desktop launch rule. The unused comtypes experiment and both staged
installer files were removed after pywin32 validation.

Host-side deterministic validation:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v xp.test_clippy_agent
=> 7 tests passed

git diff --check
=> clean
```

The tests use a fake COM object to verify single-load persistence, the exact
method surface, parameter validation, runtime animation guarding, JSON-RPC
error behavior, safe teardown, and multiple requests through one stdin
session. The same suite then passed under XP's actual Python 3.4.4 interpreter:

```text
ssh windows-xp \
  'cd /d C:\clippy && python -m unittest -v xp.test_clippy_agent'
=> 7 tests passed in 0.281s
```

The XP files were mirrored under `C:\clippy`, then the visible UTM desktop ran:

```text
cd /d C:\clippy
python -u scripts\diagnostics\smoke_clippy_controller.py
=> ANIMATIONS=<the 45 installed Clippit names>
=> VISIBLE_SMOKE_READY
=> VISIBLE_SMOKE_COMPLETE

python -u xp\clippy_agent.py
```

One persistent interactive session successfully returned JSON-RPC results for
`clippy.animations`, `clippy.show`, `clippy.move`, `clippy.play` with
`Greeting`, `clippy.think`, `clippy.speak`, and `clippy.hide`. The fabricated
animation `NotInstalled` returned `-32602` and never reached COM. Fresh visible
evidence in `docs/screenshots/clippy-global-controller.jpg` shows the real
Clippit character and his thought balloon, alongside the protocol responses.

After removing an unnecessary blocking-reader helper found during acceptance,
the clean-shutdown probe ran on the visible desktop:

```text
python -u xp\clippy_agent.py < \
  scripts\diagnostics\clippy_controller_probe.jsonl
=> id 1 returned all installed animation names
=> id 2 returned {"closing":true}
=> returned to the command prompt with no python.exe process left
```

Known limitations: this is not yet an MCP server; there is no `initialize` or
MCP tool schema. Agent request completion is not surfaced after a call queues,
and the global host still needs a defined visible-desktop launch/lifecycle
mechanism. The next architectural decision is whether the minimal MCP framing
runs directly in the Python process or wraps this JSON-RPC controller as a
child process.

## 2026-08-19 - MCP design specification for the global Clippy host

Added `docs/specs/mcp/README.md` as the implementation-ready follow-on to the
completed Phase 1/2 controller. The specification records the chosen
desktop-wide direction: real Microsoft Agent plus the installed
`CLIPPIT.ACS` character is the primary interface; Word's Office Assistant is a
separate, later capability.

The document defines the six agreed phases. Phase 3 covers MCP lifecycle
negotiation, UTF-8 newline stdio JSON-RPC, `tools/list`, and the fixed Clippy
tool set. Phase 4 limits XP automation to allowlisted read-only observations
with explicit confirmation. Phase 5 requires intentional visible handoff for
Word-native Assistant actions. Phase 6 covers XP/UTM validation and synchronized
documentation. It also records every verified XP constraint: Python 3.4.4,
32-bit pywin32 build 220, visible-session launch, SSH/non-visible Agent
behavior, the comtypes `Connected=True` hang, runtime animation guarding, and
the prohibition on arbitrary shell or unrestricted COM.

The spec leaves the required architecture decision explicit: implement MCP
directly in the Python process or put a modern MCP adapter around the current
line-delimited JSON-RPC controller. It pins no protocol version until Phase 3
implementation rechecks the supported MCP version.

Documentation-only validation:

```text
git diff --check
=> clean

rg -n "Phase 3|Phase 4|Phase 5|Phase 6|Open architecture decision" \
  docs/specs/mcp/README.md docs/ARCHITECTURE.md
=> all planned phases and the open decision present
```

No XP binaries, controller code, or runtime behavior changed in this milestone;
the remaining work is the Phase 3 implementation decision and MCP adapter or
direct-server build.

## 2026-08-19 - Phase 3 direct MCP server

Implemented the phase 3 MCP lifecycle and fixed Clippy tool surface directly
in `xp/clippy_agent.py`. The `--mcp` entry point pins protocol version
`2025-11-25`, accepts one UTF-8 JSON-RPC message per stdin line, requires
`initialize` followed by `notifications/initialized`, advertises only the
`tools` capability, and exposes the seven `clippy.*` tools from the existing
controller. Tool calls return MCP `content` plus `structuredContent`, normal
argument/animation failures are `isError` tool results, and unknown methods or
lifecycle violations are JSON-RPC errors. EOF now hides, stops, and unloads
the character before exit.

The architecture decision is direct MCP in Python rather than a modern child
process adapter. This preserves one XP COM apartment and one persistent Agent
connection; the original line JSON-RPC mode remains available for diagnostics.

Host-side validation:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v xp.test_clippy_agent
=> 11 tests passed

PYTHONDONTWRITEBYTECODE=1 python3 -m py_compile xp/clippy_agent.py
=> clean
```

The new fixtures cover initialize/version negotiation, capability discovery,
notifications, exact tool discovery, successful calls, safe tool errors,
malformed/pre-initialized requests, and clean EOF teardown using fake COM.
Known limitations remain: the real Agent calls are asynchronous and report
`queued: true` without completion/error observation, and visible XP/UTM
acceptance still must be run from the XP desktop rather than SSH.

## 2026-08-19 - Persistent Codex MCP service bridge

Added the persistent service boundary needed for Codex to reach the real
visible XP Agent process. `xp/clippy_agent.py --mcp-tcp 127.0.0.1 3211` keeps a
loopback-only MCP listener in the interactive XP desktop. The new
`xp/mcp_stdio_forward.py` is a line-oriented SSH bridge: Codex owns its stdio
child, while the bridge forwards MCP bytes to the XP listener and never creates
COM. `scripts/service/clippy_mcp_start.bat` starts the listener from Brett's
XP Startup folder.

The Codex CLI registration is now global:

```text
codex mcp add clippy -- /usr/bin/ssh -T windows-xp "python -u C:\clippy\xp\mcp_stdio_forward.py 127.0.0.1 3211"
=> Added global MCP server 'clippy'.
```

End-to-end bridge validation used that exact SSH command and returned:

```text
initialize => protocolVersion 2025-11-25, tools capability
tools/list => all seven clippy.* tools
clippy.animations => 43 runtime names from CLIPPIT.ACS
clippy.show => {"queued":true}
clippy.hide => {"queued":true}
```

The XP listener was launched from the visible UTM desktop and verified with
`netstat` on `127.0.0.1:3211`. Host and XP controller tests remain 11/11, both
Python files compile under the host interpreter, and `git diff --check` is
clean. The Codex desktop session must be restarted or opened as a new session
before the newly registered server enters its live tool inventory. The visible
XP Startup listener must be running before that session connects.

## 2026-08-21 - Codex MCP compatibility and request metadata

Fixed the Codex desktop integration in `xp/clippy_agent.py`. The server now
negotiates MCP `2025-06-18` as well as its existing `2025-11-25` version,
returns the client-requested supported version, and accepts standard `_meta`
request fields on initialize, lifecycle, `tools/list`, and `tools/call`.
Tool argument validation remains strict and does not treat `_meta` as a tool
argument.

The failure was reproduced through a captured Codex app-server request:
`tools/list` included `params._meta.progressToken`; the previous strict
validator raised the internal `InvalidParams` exception outside the
`McpError` handler, logged `Unhandled MCP server error`, and returned
JSON-RPC `-32603`.

Validation:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v xp.test_clippy_agent
=> 12 tests passed

PYTHONDONTWRITEBYTECODE=1 python3 -m py_compile xp/clippy_agent.py xp/mcp_stdio_forward.py
=> clean

Codex app-server mcpServerStatus/list through the SSH bridge
=> clippy-xp-agent with all seven clippy.* tools
```

The patched controller was copied to the XP VM and relaunched in the existing
interactive desktop session without rebooting the guest or host. The current
remaining limitation is that the live Codex task may need a fresh MCP
inventory refresh before the newly discovered tools appear in its tool list.

## 2026-09-08 - Clean MCP disconnects without tearing down Agent COM

Diagnosed a live MCP listener that remained in `CLOSE_WAIT` after its first SSH
bridge exited. `serve_mcp_socket` handled clients serially and called
`ClippyController.close()` before closing the accepted socket, so a slow Agent
COM hide/stop/unload path could prevent the listener from advancing to the next
connection. A second bridge could complete the TCP handshake but never reach
MCP handling, while another probe was refused once the one-entry backlog was
full.

The persistent TCP service now owns one `ClippyController` for its complete
visible-process lifetime. Each accepted bridge gets independent MCP lifecycle
state; EOF shuts down and closes that socket before returning to `accept()`,
without unloading the shared character. Direct stdio mode keeps its existing
EOF behavior and still hides, stops, and unloads Clippy before exit. The stdio
bridge now waits at most two seconds for final server output after Codex closes
stdin, then fully closes its socket rather than joining its receiver forever.

Regression validation on macOS and XP Python 3.4.4:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v \
  xp.test_clippy_agent xp.test_mcp_stdio_forward
=> 15 tests passed

C:\Python34\python.exe -m unittest -v \
  xp.test_clippy_agent xp.test_mcp_stdio_forward
=> 15 tests passed
```

The patched listener was launched from the visible UTM XP desktop. Seven
complete SSH bridge sessions negotiated MCP, listed all seven tools, reported
43 installed animations, and queued visible `clippy.show` and `clippy.think`
actions. After each bridge exited, `netstat -ano` showed only the listener plus
normal `TIME_WAIT` entries: no `CLOSE_WAIT` remained. Repeated post-restart
disconnect cycles kept the same listener PID 304. A fresh UTM screenshot
visibly confirmed the real Clippit character and the thought balloon `Cleanup
complete. I can accept a fresh MCP connection after the previous one closes.`

The listener remains intentionally serial and supports one active MCP client
at a time. This change guarantees prompt cleanup after that client closes; it
does not add concurrent COM clients or interrupt an Agent COM call already in
progress.

## 2026-09-08 - Add metadata-only MCP transport diagnostics

Added best-effort JSONL instrumentation to the persistent XP MCP service and
the SSH stdio bridge. `C:\clippy\ClippyMcp.log` records listener, session,
controller initialization, safe MCP method/tool routing, response write,
message-pump, error-type, and socket-close events.
`C:\clippy\ClippyMcpBridge.log` records safe stdin/TCP/stdout byte-flow and
bounded-shutdown events. Neither file records request IDs, params, tool
arguments, user text, caller-supplied animation names, or complete JSON-RPC
payloads. Logging failures are intentionally ignored so diagnostics cannot
contaminate stdout or fail MCP work.

Focused host verification:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v \
  xp.test_clippy_agent xp.test_mcp_stdio_forward
=> 19 tests passed

PYTHONDONTWRITEBYTECODE=1 python3 -m py_compile \
  xp/clippy_agent.py xp/mcp_stdio_forward.py xp/mcp_diagnostics.py
=> passed

git diff --check
=> passed
```

Copied `clippy_agent.py`, `mcp_stdio_forward.py`, `mcp_diagnostics.py`, and the
two focused test files into `C:\clippy\xp`. The first visible restart exposed
a Python 3.4 direct-script difference: a failed relative import raises
`SystemError`, not `ImportError`. Both runtime scripts now include that legacy
exception in their sibling-module fallback. The complete XP suite then passed:

```text
C:\Python34\python.exe -m unittest -v \
  xp.test_clippy_agent xp.test_mcp_stdio_forward
=> 19 tests passed on Python 3.4.4
```

Relaunched the final listener from the visible UTM desktop as PID 3016.
A standalone SSH bridge session negotiated MCP `2025-06-18`, returned all
seven tools, and returned 43 installed animations from `clippy.animations`.
The service log showed controller connect plus dispatch/response completion;
the bridge log showed all four requests crossing TCP and both response chunks
reaching stdout. Both sides recorded EOF and clean close, after which the same
listener returned to `tcp_accept_wait` with only a normal `TIME_WAIT` socket.
The same session sent an unknown method containing a unique sentinel; both
logs recorded its method as `other`, and a search confirmed that the sentinel
itself was absent from both files.

A second live session initialized, queued `clippy.show` and `clippy.think`, and
then held stdin open for 15 seconds. At ten seconds the bridge log recorded
`bridge_receive_error` with `error_type` `timeout`. This demonstrates that the
socket returned by `socket.create_connection(..., timeout=10)` retains that
timeout for normal receives: after ten idle seconds its receiver exits, while
the stdin/send side can remain alive. A later request can therefore reach XP
without any bridge thread left to forward the response, which is a strong
match for the observed Codex-side hang. Fixing that transport behavior is the
next task; this diagnostics change intentionally records rather than alters it.

The already-open Codex task still owns the pre-instrumentation SSH bridge that
was disconnected when the listener restarted. A fresh Codex app/task is
required to launch the instrumented bridge and reproduce the earlier
app-specific hanging tool call; the completed standalone acceptance proves the
new logger and the underlying XP request path, not current-task reconnection.

## 2026-09-08 - Keep idle MCP bridge sessions connected

The first `clippy.animations` call from a freshly restarted Codex task returned
all 43 installed animations in 0.5 seconds. After twelve idle seconds, the same
call failed to return within a 15-second observation window. The instrumented
bridge showed its receiver timing out at ten seconds; the next request still
crossed TCP, and the XP service dispatched it and completed its response write,
but no receive thread remained to forward that response to Codex.

Changed the stdio bridge to retain its ten-second timeout only for
`socket.create_connection`. Immediately after a successful connection it now
calls `settimeout(None)`, returning established-session receives to blocking
mode. Connection setup failure remains bounded, and the existing two-second EOF
shutdown path remains responsible for interrupting a blocked receiver.

Added a focused regression that injects the connection factory and verifies the
ten-second connect deadline followed by `settimeout(None)`. Host verification:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v \
  xp.test_clippy_agent xp.test_mcp_stdio_forward
=> 20 tests passed

git diff --check
=> passed
```

Copied the updated bridge and focused test to `C:\clippy\xp`. The complete XP
Python 3.4.4 suite passed all 20 tests. After terminating only the known stale
pre-fix bridge PID 1484, a standalone deployed bridge negotiated MCP, remained
idle from `00:56:46Z` through `00:57:00Z`, and then returned all 43 animations.
Its log recorded `receive_mode` `blocking`, the response reaching stdout after
the fourteen-second idle interval, and a clean EOF/socket close with no receive
timeout. The persistent visible listener remained PID 3016 and returned to its
next `accept()`.

## 2026-09-08 - Pump Agent COM while the MCP service is idle

The Clippy sideshow exposed a second kind of MCP idleness after the bridge
timeout fix. `clippy.show`, `clippy.move`, `clippy.play`, and `clippy.think`
all returned `queued: true`, but Clippy held the `Greeting` trumpet frame until
another MCP call arrived. The server called `PumpWaitingMessages()` only after
dispatching a request; it otherwise blocked in the active client socket's
`recv()` or the listener's `accept()`. Repeated harmless
`clippy.animations` calls acted as message-loop ticks and allowed the queued
Agent requests to advance.

Changed the persistent `--mcp-tcp` path to wait for socket readability with
`select.select()` in 50 ms intervals. Each idle interval pumps Windows messages
on the same thread that owns `Agent.Control.2`. The active client input path
does this between MCP lines, and the listener does the same while waiting for
the next client, so an action can continue after a protocol session closes.
This does not add a COM worker thread, concurrent MCP clients, tool surface, or
payload logging. Direct `--mcp` stdio remains request-driven; Codex uses the
fixed persistent TCP service.

Added regressions for an active connection that idles and later sends
`tools/list`, and for the persistent listener pumping before its next
`accept()`. Host validation passed:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest -v \
  xp.test_clippy_agent xp.test_mcp_stdio_forward
=> 22 tests passed

PYTHONDONTWRITEBYTECODE=1 python3 -m py_compile \
  xp/clippy_agent.py xp/mcp_stdio_forward.py xp/mcp_diagnostics.py
=> passed

git diff --check
=> passed
```

Copied `xp/clippy_agent.py` and `xp/test_clippy_agent.py` to `C:\clippy\xp`.
The same 22 tests passed under XP's Python 3.4.4. Terminated only the known old
Clippy owner, Agent server, and stale bridge processes, then relaunched
`scripts\service\clippy_mcp_start.bat` from the visible UTM desktop as PID
3652.

Live acceptance opened one standalone MCP bridge, queued `show`, `move`,
`Greeting`, and `think`, and then sent zero MCP requests for twelve seconds. A
fresh UTM capture during that silent interval showed Clippy advance from the
trumpet frame into the Think animation. The ordinary `tools/list` request sent
after the pause returned successfully. After bridge EOF, a later fresh capture
showed Clippy continuing to animate while the service waited in `accept()`.
A second balloon-only session also held the connection silent for twelve
seconds and visibly advanced Clippy at the requested in-bounds position.

`C:\clippy\ClippyMcp.log` records no MCP requests between `01:36:18Z` and EOF
at `01:36:30Z`, followed by clean client close and `tcp_accept_wait`. `netstat`
showed PID 3652 listening on `127.0.0.1:3211`, two normal `TIME_WAIT` entries,
and no `CLOSE_WAIT`. High-frequency idle pumps are intentionally not logged so
metadata diagnostics do not grow by twenty records per second. Agent actions
still report only that they were queued; request completion/error observation
remains separate future work.

After adding the final guard that skips idle pumping before the controller has
connected, copied the controller to XP again and repeated the 22-test Python
3.4.4 suite. Relaunched that exact final source from UTM as PID 2812. A fresh
standalone probe negotiated MCP `2025-06-18`, listed all seven tools and 43
installed animations, then closed cleanly; `netstat` showed PID 2812 listening
with one normal `TIME_WAIT` connection and no `CLOSE_WAIT`.
