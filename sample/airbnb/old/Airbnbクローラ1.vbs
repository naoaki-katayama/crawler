'http://www.ken3.org/vba/backno/vba155.html'
Sub Airbnbクローラ1 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String
  Dim htmlall As String

  i = 14953606
  URLstr = "https://www.airbnb.jp/rooms/" & i
  Call ieView(objIE, URLstr)
  URLrow = 1
  Call tagClick(objIE,"input","datespan-checkin")
  Application.Wait (DateAdd("s", 5, Now))
  Call tagClick(objIE,"a","ui-datepicker-next icon icon-chevron-right ui-corner-all")
  Application.Wait (DateAdd("s", 5, Now))
  Call tagMouseover(objIE,"a","ui-state-default")
  Application.Wait (DateAdd("s", 5, Now))

  htmlall =  objIE.document.all(0).outerHTML
  'Call Cellに記入(htmlall)
  'Debug.Print htmlall
  objIE.Quit
End Sub

Sub Cellに記入(htmlall As String)
  Dim inputColumn As Long
  inputColumn = Cells(1,10000).End(xlToLeft).Column + 1
  htmlLine = Split(htmlall , ">")
  For j = LBound(htmlLine) To UBound(htmlLine)
      Cells(j + 1, inputColumn) = j
      Cells(j + 1, inputColumn + 1) = htmlLine(j) & ">"
      DoEvents
  Next j
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
'==================================================='
Sub tagMouseover(objIE As InternetExplorer, _
             tagName As String, _
             tagStr As String)

  'タグをマウスオーバー
  For Each objTag In objIE.document.getElementsByTagName(tagName)
  Stop
  Debug.Print objTag.outerHTML
    If InStr(objTag.outerHTML, tagStr) > 0 Then
      objTag.onmouseover
      Debug.Print objTag'texttext
      Call ieCheck(objIE)
      Exit For
    End If
  Next
End Sub
