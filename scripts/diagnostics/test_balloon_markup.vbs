Option Explicit

Dim word, balloon

Set word = CreateObject("Word.Application")
word.Visible = True
word.Activate
WScript.Sleep 1500

Set balloon = word.Assistant.NewBalloon
With balloon
    .Heading = "Native balloon formatting"
    .Text = "Plain text, {ul}underlined emphasis{ul 0}, " & _
        "{cf 4}`blue code`{cf 0}, and {cf 2}green text{cf 0}." & vbCrLf & vbCrLf & _
        "{ul}{cf 4}Section heading{cf 0}{ul 0}" & vbCrLf & _
        "A compact second paragraph."
    .BalloonType = 1
    .Labels(1).Text = "First native bullet"
    .Labels(2).Text = "Second native bullet"
    .Button = 1
    .Show
End With
