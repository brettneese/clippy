Option Explicit

Dim agent, clippy, animationName

Set agent = CreateObject("Agent.Control.2")
agent.Connected = True
agent.Characters.Load "Clippy", "C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS"
Set clippy = agent.Characters("Clippy")

For Each animationName In clippy.AnimationNames
  WScript.Echo animationName
Next

agent.Characters.Unload "Clippy"
