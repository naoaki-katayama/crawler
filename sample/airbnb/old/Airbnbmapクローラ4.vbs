Sub Airbnbmapクローラ4 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String
  Dim htmlall As String
  Dim SearchStr As String
  Dim pageCount As Long
  Dim ws(1 to 2) As Worksheet
  Set ws(1) = Worksheets("大阪市")
  Set ws(2) = Worksheets("変数定義")
  Dim ieCount As Integer
  ieCount = 0
  Dim i As Long

  For i = 1 to ws(2).Cells(1,1).End(xlDown).Row
    ws(1).Cells(1,1) = i
    URLstr = "https://www.airbnb.jp/s/%E5%A4%A7%E9%98%AA%E5%B8%82?page=1&source=map&airbnb_plus_only=false" _
      & "&sw_lat=" & ws(2).Cells(i,3) _
      & "&sw_lng=" & ws(2).Cells(i,4) _
      & "&ne_lat=" & ws(2).Cells(i,5) _
      & "&ne_lng=" & ws(2).Cells(i,6) _
      & "&search_by_map=true"
    Debug.Print "南北：" & ws(2).Cells(i,1) & ",東西：" & ws(2).Cells(i,2) & "," & URLstr
    If ieCount = 0 Then
      Call ieView(objIE, URLstr)
      ieCount = 1
    Else
      Call ieNavi(objIE, URLstr)
    End If

    Application.Wait (DateAdd("s", 3, Now))

    htmlall =  objIE.document.all(0).outerHTML
    pageCount = 2
    Do
      'スクレイピング実施
      Call Inputhtml(objIE,ws(2).Cells(i,1).Value,ws(2).Cells(i,2).Value)
      SearchStr = "page=" & pageCount
      Debug.Print "pageCount:" & pageCount
      Call tagClick2(objIE,"a",SearchStr,pageCount)
      Application.Wait (DateAdd("s", 3, Now))

      pageCount = pageCount + 1
    Loop while pageCount > 1
    ws(1).Cells.WrapText = False

    ActiveWorkbook.Save
  Next i
  objIE.Quit
End Sub
'==================================================='
Sub Inputhtml(objIE As InternetExplorer, _
             snNum As Long, _
             weNum As Long)

  Dim ws(1) As Worksheet
  Set ws(1) = Worksheets("大阪市")
  Dim InputRow As Long
  Dim str As String
  InputRow = ws(1).Cells(1,4).End(xlDown).Row + 1

  For Each objTag In objIE.document.getElementsByTagName("span")
    str = objTag.outerHTML
    If InStr(objTag.outerHTML, "span tabindex") > 0 Then
      ws(1).Cells(InputRow,1) = ws(1).Cells(InputRow,2).Row - 2
      ws(1).Cells(InputRow,2) = snNum
      ws(1).Cells(InputRow,3) = weNum
      ws(1).Cells(InputRow,4) = Mid(str,InStr(str,"data-hosting_id") + 17,InStr(str,"data-address") - InStr(str,"data-hosting_id") - 19) '部屋ID（URL)
      ws(1).Cells(InputRow,6) = Mid(str,InStr(str,"data-star_rating") + 18,InStr(str,"data-room_type") - InStr(str,"data-star_rating") - 20) '評価
      ws(1).Cells(InputRow,7) = Mid(str,InStr(str,"data-review_count") + 19,InStr(str,"data-hosting_id") - InStr(str,"data-review_count") - 21) 'レビュー数
      ws(1).Cells(InputRow,8) = Mid(str,InStr(str,"data-room_type") + 16,InStr(str,"data-review_count") - InStr(str,"data-room_type") - 18) '貸タイプ
      ws(1).Cells(InputRow,9) = Mid(str,InStr(str,"data-property_type_name") + 25,InStr(str,"data-host_img") - InStr(str,"data-property_type_name") - 27) '建物タイプ
      ws(1).Cells(InputRow,10) = Mid(str,InStr(str,"data-bedrooms_string") + 22,InStr(str,"data-person_capacity_string") - InStr(str,"data-bedrooms_string") - 24) 'ベッドルーム数
      ws(1).Cells(InputRow,11) = Mid(str,InStr(str,"data-person_capacity_string") + 29,InStr(str,"data-property_type_name") - InStr(str,"data-person_capacity_string") - 31) '定員
      ws(1).Cells(InputRow,12) = Mid(str,InStr(str,"data-address") + 14,InStr(str,"data-name") - InStr(str,"data-address") - 16) '住所
      ws(1).Cells(InputRow,13) = Mid(str,InStr(str,"data-name") + 11,InStr(str,"data-img") - InStr(str,"data-name") - 13) '部屋タイトル
      ws(1).Cells(InputRow,14) = Mid(str,InStr(str,"data-host_id") + 14,InStr(str,"data-star_rating") - InStr(str,"data-host_id") - 16) 'ホストID(URL)
      ws(1).Cells(InputRow,15) = Mid(str,InStr(str,"data-img") + 10,InStr(str,"><") - InStr(str,"data-img") - 31) '部屋画像
      ws(1).Cells(InputRow,16) = Mid(str,InStr(str,"data-host_img") + 15,InStr(str,"data-host_id") - InStr(str,"data-host_img") - 43) 'ホスト画像

      InputRow = InputRow + 1
    End If
    DoEvents
  Next

  InputRow = ws(1).Cells(InputRow,5).End(xlUp).Row + 1
  For Each objTag In objIE.document.getElementsByTagName("span")
    If InStr(objTag.outerHTML, "price-amount") > 0 and InStr(objTag.outerHTML,"/sup") <> 0 Then

      ws(1).Cells(InputRow,5) = Mid(objTag.innerText,2)
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
