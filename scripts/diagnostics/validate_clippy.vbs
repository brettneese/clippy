Option Explicit

Dim agent, clippy, anim, names
Set agent = CreateObject("Agent.Control.2")
agent.Connected = True
agent.Characters.Load "Clippy", "C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS"
Set clippy = agent.Characters("Clippy")

names = ""
For Each anim In clippy.AnimationNames
  names = names & anim & vbCrLf
Next

WScript.Echo names
agent.Characters.Unload "Clippy"
Set clippy = Nothing
Set agent = Nothing
