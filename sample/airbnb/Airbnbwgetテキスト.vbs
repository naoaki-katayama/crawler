Sub Airbnbroomwget()

Dim ws(1) As Worksheet
Set ws(1) = Worksheets("マクロメニュー")

Dim rc As Long
Dim htmlStr As String
Dim cmd As String
'For i = 1 to 20000000
For i = 9798417 to 9798421
htmlStr = "https://www.airbnb.jp/rooms/" & i

cmd = "wget " & htmlStr & " -P D:\airb" & " -O airb-room" & i & ".txt"
rc = Shell(cmd, vbNormalFocus)
If rc = 0 Then
 MsgBox "起動に失敗しました"
End If
Application.Wait (DateAdd("s", 1, Now))

Next i

End Sub
