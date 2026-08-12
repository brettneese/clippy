# ClippyProbe development log

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
