Option Explicit

Const Endpoint = "http://10.0.2.2:3210/message"

Dim request
Set request = CreateObject("MSXML2.ServerXMLHTTP.3.0")
request.setTimeouts 3000, 3000, 3000, 3000
request.open "POST", Endpoint, False
request.setRequestHeader "Content-Type", "application/json; charset=utf-8"
request.send "{""text"":""hello from Windows XP""}"

WScript.Echo "HTTP " & request.status
WScript.Echo request.responseText

If request.status <> 200 Then
  WScript.Quit 1
End If
