Option Explicit

Dim agent, clippy, shell, fso, tempDir, statusPath

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
tempDir = shell.ExpandEnvironmentStrings("%TEMP%")
statusPath = tempDir & "\clippy-stage-status.txt"

Sub SetStatus(ByVal text)
  Dim statusFile
  Set statusFile = fso.CreateTextFile(statusPath, True)
  statusFile.WriteLine text
  statusFile.Close
End Sub

Sub Pause(ByVal milliseconds)
  WScript.Sleep milliseconds
End Sub

Sub Play(ByVal animationName, ByVal milliseconds)
  clippy.StopAll
  clippy.Play animationName
  Pause milliseconds
  clippy.StopAll
  Pause 300
End Sub

Sub Say(ByVal text, ByVal milliseconds)
  clippy.StopAll
  clippy.Speak text
  Pause milliseconds
  clippy.StopAll
  Pause 500
End Sub

Sub Think(ByVal text, ByVal milliseconds)
  clippy.StopAll
  clippy.Think text
  Pause milliseconds
  clippy.StopAll
  Pause 500
End Sub

SetStatus "Starting"

Set agent = CreateObject("Agent.Control.2")
agent.Connected = True
agent.Characters.Load "Clippy", "C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS"
Set clippy = agent.Characters("Clippy")
clippy.Balloon.Style = 7

clippy.MoveTo 690, 430
clippy.Show
Pause 1800

SetStatus "Opening number"
Play "Greeting", 2500
Say "Ladies and gentlemen, welcome to the one paperclip show!", 8500
Play "GetAttention", 2200
Think "Tonight: drama, suspense, and absolutely no useful formatting advice.", 9000

SetStatus "Act One"
clippy.MoveTo 80, 380
Pause 2600
Play "GetArtsy", 3000
Think "ACT ONE: A humble paperclip dreams of becoming... a slightly more dramatic paperclip.", 10000
Play "Writing", 3500
Say "I wrote my entire memoir in Microsoft Word. It is mostly tooltips.", 9000

SetStatus "Act Two"
clippy.MoveTo 720, 170
Pause 2600
Play "Searching", 3500
Think "ACT TWO: Our hero searches the desktop for meaning... and the missing Start button.", 10000
Play "CheckingSomething", 3200
Say "Aha! The meaning of life is forty two. The meaning of Windows XP is: please wait.", 10000

SetStatus "Finale"
clippy.MoveTo 390, 360
Pause 2600
Play "GetTechy", 3200
Think "FINALE: If at first you do not succeed, animate confidently and call it a feature!", 10000
Play "Congratulate", 3600
Say "You have been a wonderful audience. Please save your applause as a dot doc file.", 9500
Play "Wave", 3300
Think "Good night! May all your documents be recovered automatically.", 9500
Play "GoodBye", 3000

SetStatus "Complete"
clippy.Hide
Pause 1800
agent.Characters.Unload "Clippy"
