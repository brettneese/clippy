# Windows XP / Clippy workflow

## Project documentation and commits

- Read `docs/ARCHITECTURE.md` before making architectural or implementation changes. Treat it as the source of truth for component boundaries, milestones, and the macOS/XP split; update it when those decisions change.
- Append a dated entry to `docs/DEVLOG.md` for meaningful implementation, reverse-engineering, build, or validation work. Record what changed, the commands or interaction used to verify it, the observed result, and any remaining limitation.
- Before finishing a coherent unit of work, run the relevant XP build or validation, review the complete diff, and commit the task's source and documentation changes with a descriptive message. Do not commit generated binaries, runtime logs, temporary files, or unrelated user changes.

- Use `ssh windows-xp` to prepare files and run diagnostics. Launch GUI programs from the visible XP desktop with UTM Computer Use and **Windows Key->R**; SSH-launched processes may run on an invisible desktop. Do not use computer use when ssh will do.
- Store short `.vbs` or `.bat` files under `%TEMP%` or `My Documents\Clippy Demo`, and send CRLF (`\r\n`) to the Windows SSH PTY. On this VM, Brett's `%TEMP%` expands to `C:\DOCUME~1\Brett\LOCALS~1\Temp`; it is not `C:\WINDOWS\TEMP`. Check with `echo %TEMP%`, and copy the script into the resolved profile directory before launching it. An `scp` target such as `C:/WINDOWS/TEMP/show.vbs` does not satisfy a later `%TEMP%\show.vbs` launch.
- Before a demo, check for stale processes with `tasklist | findstr /i "wscript winword agentsvr"`. Close only known demo instances; an SSH-launched Word instance can capture later launches.
- Clippit's installed character file is `C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS`, not the usual `C:\WINDOWS\msagent\chars` location. Before using animation names, load that file with `Agent.Control.2`, enumerate `Character.AnimationNames`, and use only names returned by the installed character. Do not assume an animation exists from examples found elsewhere.
- The animation names returned by the currently installed Clippit character are:

  ```text
  GestureLeft
  GestureDown
  GestureRight
  GestureUp
  IdleFingerTap
  IdleSideToSide
  IdleEyeBrowRaise
  GetArtsy
  LookDownRight
  LookDown
  LookDownLeft
  LookRight
  LookLeft
  Hearing_1
  LookUpRight
  LookUp
  LookUpLeft
  GetAttention
  Wave
  Save
  Congratulate
  Processing
  GoodBye
  Print
  CheckingSomething
  EmptyTrash
  IdleRopePile
  IdleSnooze
  IdleHeadScratch
  SendMail
  Searching
  Thinking
  Writing
  Explain
  RestPose
  Show
  Hide
  GetTechy
  GetWizardy
  Alert
  IdleAtom
  Greeting
  Idle1_1
  ```

- DO NOT USE CAPTURE INPUT in UTM
- Launch scripts with a short command such as `wscript.exe "%TEMP%\show.vbs"`, then verify success with fresh UTM screenshots—not `tasklist` alone. Capture at least one screenshot showing Clippy and another showing a balloon or a later animation; `wscript.exe` plus `agentsvr.exe` only confirms that the controller started.
- For interactive Office-style balloons, prefer Word's native `Assistant.NewBalloon` API. Activate Word and allow a short startup delay before calling `Balloon.Show`.
- `Agent.Character.Speak` is asynchronous and a screenshot may catch its balloon before text is painted or after it clears. Allow enough time before visual verification. When visible text matters more than speech, use `Character.Think` or Word's `Assistant.NewBalloon` and verify the text in a fresh screenshot.
- Avoid simultaneous Word or Clippy controllers. Use Clippy as the interface and VBScript/COM automation as the mechanism for approved actions.
