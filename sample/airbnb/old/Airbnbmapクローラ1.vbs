Sub Airbnbmapクローラ1 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String
  Dim htmlall As String
  Dim SearchStr As String

'  i = 4756460
  URLstr = "https://www.airbnb.jp/s/%E5%A4%A7%E9%98%AA%E5%B8%82?page=1&source=map&airbnb_plus_only=false&sw_lat=34.6637733136886&sw_lng=135.470922729394&ne_lat=34.6711533136886&ne_lng=135.478732729394&search_by_map=true"
  Call ieView(objIE, URLstr)
  URLrow = 1
'  Call tagClick(objIE,"input","datespan-checkin")
'  Application.Wait (DateAdd("s", 5, Now))
'  Call tagClick(objIE,"a","ui-datepicker-next icon icon-chevron-right ui-corner-all")
'  Application.Wait (DateAdd("s", 3, Now))
'  SearchStr = Cells(1, 1)
'  Call tagClick(objIE,"a",SearchStr)

  htmlall =  objIE.document.all(0).outerHTML
  Call Cellに記入(htmlall)
  htmlall =  objIE.document.all(0).innerText
  Call Cellに記入(htmlall)
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
  Debug.Print objTag.outerHTML'ｔｅｓｔ
    If InStr(objTag.outerHTML, tagStr) > 0 Then
      objTag.Click
      Call ieCheck(objIE)
      Exit For
    End If
  Next
End Sub
