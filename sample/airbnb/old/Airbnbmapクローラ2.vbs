Sub Airbnbmapクローラ2 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String
  Dim htmlall As String
  Dim SearchStr As String
  Dim pageCount As Long
  Dim ws(1 to 2) As Worksheet
  Set ws(1) = Worksheets("大阪市")
  Set ws(2) = Worksheets("変数定義")

  Dim i As Long
  'Dim StartRangte As Range

  For i = 199 to 199'ws(2).Cells(1,1).End(xlDown)
    URLstr = "https://www.airbnb.jp/s/%E5%A4%A7%E9%98%AA%E5%B8%82?page=1&source=map&airbnb_plus_only=false" _
      & "&sw_lat=" & ws(2).Cells(i,3) _
      & "&sw_lng=" & ws(2).Cells(i,4) _
      & "&ne_lat=" & ws(2).Cells(i,5) _
      & "&ne_lng=" & ws(2).Cells(i,6) _
      & "&search_by_map=true"
      Debug.Print "南北：" & ws(2).Cells(i,1) & ",東西：" & ws(2).Cells(i,2) & "," & URLstr
      Call ieView(objIE, URLstr)

'  URLrow = 1
'  Call tagClick(objIE,"input","da'tespan-checkin")
'  Application.Wait (DateAdd("s", 5, Now))
'  Call tagClick(objIE,"a","ui-datepicker-next icon icon-chevron-right ui-corner-all")
'  Application.Wait (DateAdd("s", 3, Now))
'  SearchStr = Cells(1, 1)
'  Call tagClick(objIE,"a",SearchStr)

  htmlall =  objIE.document.all(0).outerHTML
  pageCount = 1
  Do
    'スクレイピング実施
    Call Inputhtml(objIE,ws(2).Cells(i,1).Value,ws(2).Cells(i,2).Value)
    SearchStr = "page=" & pageCount
    Debug.Print "pageCount:" & pageCount
    Call tagClick2(objIE,"a",SearchStr,pageCount)

    pageCount = pageCount + 1
  Loop while pageCount > 1
  ws(1).Cells.WrapText = False

  'Call Cellに記入(htmlall)
  'Call Cellに記入2(htmlall)
  'Debug.Print htmlall
  Stop
  objIE.Quit
  Next i
End Sub
'==================================================='
Sub Inputhtml(objIE As InternetExplorer, _
             snNum As Long, _
             weNum As Long)

  Dim ws(1) As Worksheet
  Set ws(1) = Worksheets("大阪市")
  Dim InputRow As Long
  InputRow = ws(1).Cells(1,3).End(xlDown).Row + 1

  For Each objTag In objIE.document.getElementsByTagName("a")
    If InStr(objTag.outerHTML, "/rooms/") > 0 and InStr(objTag.outerHTML, "<a href=") Then
      ws(1).Cells(InputRow,2) = ws(1).Cells(InputRow,2).Row - 2
      ws(1).Cells(InputRow,3) = snNum
      ws(1).Cells(InputRow,4) = weNum
      ws(1).Cells(InputRow,5) = Mid(objTag.outerHTML,17,InStr(objTag.outerHTML,"target") - 19)
      InputRow = InputRow + 1
    End If
    DoEvents
  Next

  InputRow = ws(1).Cells(InputRow,6).End(xlUp).Row + 1
  For Each objTag In objIE.document.getElementsByTagName("span")
    If InStr(objTag.outerHTML, "price-amount") > 0 Then
      ws(1).Cells(InputRow,6) = objTag.innerText
      InputRow = InputRow + 1
    End If
    DoEvents
  Next

  InputRow = ws(1).Cells(InputRow,7).End(xlUp).Row + 1
  For Each objTag In objIE.document.getElementsByTagName("h3")
    If InStr(objTag.outerHTML, "title") > 0 Then
      ws(1).Cells(InputRow,7) = Mid(objTag.outerHTML,12,InStr(objTag.outerHTML,"class=") - 14)
      InputRow = InputRow + 1
    End If
    DoEvents
  Next


End Sub

'==================================================='

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

Sub Cellに記入2(htmlall As String)
  Dim inputColumn As Long
  Dim ws(3) As Worksheet
  Set ws(3) = Worksheets("htmlテスト")
  ws(3).Cells.Clear
  htmlLine = Split(htmlall , ">")
  For j = LBound(htmlLine) To UBound(htmlLine)
      ws(3).Cells(j + 1, 1) = j
      ws(3).Cells(j + 1, 2) = htmlLine(j) & ">"
      DoEvents
  Next j
  ws(3).Cells.WrapText = False
End Sub

'==================================================='
Sub tagClick(objIE As InternetExplorer, _
             tagName As String, _
             tagStr As String)

  'タグをクリック
  For Each objTag In objIE.document.getElementsByTagName(tagName)
  'Debug.Print objTag.outerHTML'ｔｅｓｔ
    If InStr(objTag.outerHTML, tagStr) > 0 Then
      objTag.Click
      Call ieCheck(objIE)
      Exit For
    End If
  Next
End Sub

'==================================================='
Sub tagClick2(objIE As InternetExplorer, _
             tagName As String, _
             tagStr As String, _
             pageCount As Long)
  Dim CountCheck As Integer
  CountCheck = 1
  'タグをクリック
  For Each objTag In objIE.document.getElementsByTagName(tagName)
  'Debug.Print objTag.outerHTML'ｔｅｓｔ
    If InStr(objTag.outerHTML, tagStr) > 0 Then
      objTag.Click
      Call ieCheck(objIE)
      CountCheck = 2
      Exit For
    End If
  Next
  If CountCheck = 1 Then
    pageCount = 0
  End If
End Sub
