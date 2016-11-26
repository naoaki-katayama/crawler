Sub Airbnbroomクローラ1 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String
  Dim htmlall As String
  Dim SearchStr As String
  Dim pageCount As Long
  Dim ws(1 to 2) As Worksheet
  Set ws(1) = Worksheets("room")

      Set ws(2) = Worksheets("テスト")

  Dim ieCount As Integer
  ieCount = 0
  Dim i As Long

  For i = 1 to 1'3 to ws(3).Cells(2,2).End(xlDown).Row - 2
    URLstr = "https://www.airbnb.jp/rooms/" & ws(1).Cells(i,3)
    Debug.Print i - 2
    If ieCount = 0 Then
      Call ieView(objIE, URLstr)
      ieCount = 1
    Else
      Call ieNavi(objIE, URLstr)
    End If

    'Application.Wait (DateAdd("s", 3, Now))

    j = 1
    For Each objTag In objIE.document.getElementsByTagName("p")
      'str = objTag.outerHTML
''      If InStr(objTag.outerHTML, "span tabindex") > 0 Then
        ws(2).Cells(j,1) = objTag.innerText
        j = j + 1
''      End If
      DoEvents
    Next

    ws(2).Cells.WrapText = False

    ActiveWorkbook.Save
  Next i
  objIE.Quit
End Sub


Sub Airbnb_roomクローラ1 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String

  URLstr = "https://www.airbnb.jp/s/大阪市"
  Call ieView(objIE, URLstr)
  URLrow = 1
  For j = 0 to objIE.document.Links.Length - 1
    LinkName = objIE.document.Links(j).href
    If InStr(LinkName, "/rooms/") <> 0 Then
      Cells(URLrow, 1) = LinkName
      URLrow = URLrow + 1
    End If
  Next j

End Sub
