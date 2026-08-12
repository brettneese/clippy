Option Explicit

Dim agent, clippy

Set agent = CreateObject("Agent.Control.2")
agent.Connected = True
agent.Characters.Load "Clippy", "C:\Program Files\Microsoft Office\Office10\CLIPPIT.ACS"
Set clippy = agent.Characters("Clippy")

clippy.Balloon.Style = 7
clippy.MoveTo 650, 420
clippy.Show
WScript.Sleep 1800
clippy.Play "Greeting"
WScript.Sleep 2500
clippy.Think "Hello! I was opened from the Windows Run dialog."
WScript.Sleep 15000

clippy.StopAll
clippy.Hide
WScript.Sleep 1000
agent.Characters.Unload "Clippy"
