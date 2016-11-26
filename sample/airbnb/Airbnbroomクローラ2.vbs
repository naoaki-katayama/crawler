Sub Airbnbroomクローラ2 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String
  Dim ws(1 to 2) As Worksheet
  Set ws(1) = Worksheets("room")
  Set ws(2) = Worksheets("変数定義")
  Dim i As Long
  Dim ieCount As Integer


  For i = 2 to 2'ws(2).Cellsw(1,2).End(xlDown).Row
    URLstr = "https://www.airbnb.jp/rooms/" & ws(2).Cells(i,1)
    If ieCount = 0 Then
      Call ieView(objIE, URLstr)
      ieCount = 1'Call IE
    Else
      Call ieNavi(objIE, URLstr)
    End If
    Call tagClick(objIE,"button","expandable-trigger-more btn-link btn-link--bold")
    ws(1).Cells(ws(1).Cells(2,1).End(xlDown).Row + 1,1) = ws(2).Cells(i,1)
    Call AirbnbRoomScraping(objIE)
  Next i
End Sub

'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Sub AirbnbRoomScraping(objIE As InternetExplorer)
Dim str As String
  For Each objTag In objIE.document.getElementsByTagName("div")
    str = objTag.outerHTML
    str2 = objTag.innerText
    Call SpacePrice(str,str2)
    Call amenity(str,str2)
  Next
End Sub
'==================================================='
Sub tagClick(objIE As InternetExplorer, _
             tagName As String, _
             tagStr As String)

  'タグをクリック
  For Each objTag In objIE.document.getElementsByTagName(tagName)
    If InStr(objTag.outerHTML, tagStr) > 0 Then
      objTag.Click
      Call ieCheck(objIE)
      Exit For
    End If
  Next
End Sub

'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Sub amenity(str As String, _
                str2 As String)

  'Dim AlreadyCheck(11 to ???) As Boolean
  Dim ws As Wroksheet
  Set ws = Worksheets("room")
  Dim inputRow = ws.Cells(1,2).End(xlDown).Row + 1

  For i = 11 to ???
    AlreadyCheck(i) = False
  Next i

  If InStr(str, "敷地内無料駐車場") > 0 and InStr(str, "キッチン") > 0 and InStr(str, "インターネット") > 0 _
    InStr(str, "ジャクージ&高級風呂") <> 0 and InStr(str, "ビル内にエレベーターあり") <> 0 Then

  ElseIf InStr(str, "ジャクージ&高級風呂") > 0 and InStr(str, "ビル内にエレベーターあり") > 0 _
    InStr(str, "敷地内無料駐車場") <> 0 and InStr(str, "キッチン") <> 0 and InStr(str, "インターネット") <> 0 Then
  End If


'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Sub SpacePrice(str As String, _
                str2 As String)

  Dim AlreadyCheck(2 to 18) As Boolean
  Dim ws As Wroksheet
  Set ws = Worksheets("room")
  Dim inputRow = ws.Cells(1,2).End(xlDown).Row + 1

  For i = 2 to 18
    AlreadyCheck(i) = False
  Next i

  If InStr(str, "div class=") > 0 Then
  Else
    If InStr(str, "span data-reactid=") > 0 Then
      If AlreadyCheck(2) = False and InStr(str, "収容人数") > 0 Then
        ws.Cells(InputRow,2) = str2
        AlreadyCheck(2) = True
      ElseIf AlreadyCheck(3) = False and InStr(str, "バスルーム") > 0 Then
        ws.Cells(InputRow,3) = str2
        AlreadyCheck(3) = True
      ElseIf AlreadyCheck(4) = False and InStr(str, "ベッドルーム") > 0 Then
        ws.Cells(InputRow,4) = str2
        AlreadyCheck(4) = True
      ElseIf AlreadyCheck(5) = False and InStr(str, "ベッド数") > 0 Then
        ws.Cells(InputRow,5) = str2
        AlreadyCheck(5) = True
      ElseIf AlreadyCheck(6) = False and InStr(str, "ペット所有") > 0 Then
        ws.Cells(InputRow,6) = str2
        AlreadyCheck(6) = True
      ElseIf AlreadyCheck(7) = False and InStr(str, "チェックイン") > 0 Then
        ws.Cells(InputRow,7) = str2
        AlreadyCheck(7) = True
      ElseIf AlreadyCheck(8) = False and InStr(str, "チェックアウト") > 0 Then
        ws.Cells(InputRow,8) = str2
        AlreadyCheck(8) = True
      ElseIf AlreadyCheck(9) = False and InStr(str, "建物タイプ") > 0 Then
        ws.Cells(InputRow,9) = str2
        AlreadyCheck(9) = True
      ElseIf AlreadyCheck(10) = False and InStr(str, "部屋タイプ") > 0 Then
        ws.Cells(InputRow,10) = str2
        AlreadyCheck(10) = True
      ElseIf AlreadyCheck(11) = False and InStr(str, "追加人数の料金") > 0 Then
        ws.Cells(InputRow,11) = str2
        AlreadyCheck(11) = True
      ElseIf AlreadyCheck(12) = False and InStr(str, "清掃料金") > 0 Then
        ws.Cells(InputRow,12) = str2
        AlreadyCheck(12) = True
      ElseIf AlreadyCheck(13) = False and InStr(str, "保証金") > 0 Then
        ws.Cells(InputRow,13) = str2
        AlreadyCheck(13) = True
      ElseIf AlreadyCheck(14) = False and InStr(str, "週の割引率") > 0 Then
        ws.Cells(InputRow,14) = str2
        AlreadyCheck(14) = True
      ElseIf AlreadyCheck(15) = False and InStr(str, "月額割引率") > 0 Then
        ws.Cells(InputRow,15) = str2
        AlreadyCheck(15) = True
      ElseIf AlreadyCheck(16) = False and InStr(str, "キャンセル") > 0 Then
        ws.Cells(InputRow,16) = str2
        AlreadyCheck(16) = True
      ElseIf AlreadyCheck(17) = False and InStr(str, "週末料金") > 0 Then
        ws.Cells(InputRow,17) = str2
        AlreadyCheck(17) = True
      ElseIf AlreadyCheck(18) = False and InStr(str, "返答時間:") > 0 Then
        ws.Cells(InputRow,18) = str2
        AlreadyCheck(18) = True
      End If
      DoEvents
      'Debug.Print str
    End If
  End If
End Sub
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
