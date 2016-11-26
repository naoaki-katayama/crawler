Sub Airbnbwgetしたテキスト変換()

Dim ws(1 to 2) As Worksheet
Set ws(1) = Worksheets("マクロメニュー")
Set ws(2) = Worksheets("マクロメニュー")

Dim buf As String
Dim wgettext As Variant
Open "C:\Users\尚亮\Dropbox\MKP\6.Airbnb\クローリング\airbnb10112039.txt" For Input As #1
  Line Input #1, buf
  wgettext = Split(buf,">")
  Debug.Print UBound(wgettext)
  Row = 1
  For i = 0 to UBound(wgettext) - 1
    'Debug.Print wgettext(i)
    If Instr(wgettext(i), "href") <> 0 Then
      Cells(Row, 1) = wgettext(i) & ">"
      Row = Row + 1
    End If
    Doevents
  Next i
 Close #1

End Sub
