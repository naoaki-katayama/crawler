Sub Airbnbmapクローラ4test ()

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

  For i = 200 to 200'1 to ws(2).Cells(1,1).End(xlDown).Row
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
      'Call Inputhtml(objIE,ws(2).Cells(i,1).Value,ws(2).Cells(i,2).Value)
      Call   Airbnbmaptextscraping(objIE,ws(2))
      SearchStr = "page=" & pageCount
      Debug.Print "pageCount:" & pageCount
      Stop
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

Sub Airbnbmaptextscraping(objIE As InternetExplorer, _
                          ws(2) As Worksheet)
  Dim str As String
  Dim str2 As Variant
  Dim str3(1 to 26) As String
  Dim detail20(1 to 4) As String
  Dim detail21(1 to 4) As String
  Dim detail25(1 to 11) As String
  Dim i As Long

  For Each objTag In objIE.document.getElementsByTagName("script")
    str = objTag.outerHTML
    If InStr(objTag.outerHTML, "data-hypernova-key") > 0 Then
      str = Mid(str,Instr(str,"""listing"""),Instr(str,"""metadata""") - Instr(str,"""listing"""))
      str = Mid(str,1,Instr(str,"}]"))
      str2 = split(str,"},{")
      'Debug.Print "UBound(str2):" & UBound(str2)
      For i = LBound(str2) To UBound(str2)
        Stop
        str3(1) = cutText(str2(i),2,2,"bedrooms","beds")
        str3(2) = cutText(str2(i),2,2,"beds","airbnb_plus_enabled")
        str3(3) = cutText(str2(i),2,2,"airbnb_plus_enabled","extra_host_languages")
        str3(4) = cutText(str2(i),3,3,"extra_host_languages","id")
        str3(5) = cutText(str2(i),2,2,"id","instant_bookable")
        str3(6) = cutText(str2(i),2,2,"instant_bookable","is_business_travel_ready")
        str3(7) = cutText(str2(i),2,2,"is_business_travel_ready","is_new_listing")
        str3(8) = cutText(str2(i),2,2,"is_new_listing","lat")
        str3(9) = cutText(str2(i),2,2,"lat","lng")
        str3(10) = cutText(str2(i),2,2,"lng","name")
        str3(11) = cutText(str2(i),3,3,"name","person_capacity")
        str3(12) = cutText(str2(i),2,2,"person_capacity","picture_count")
        str3(13) = cutText(str2(i),2,2,"picture_count","picture_url")
        str3(14) = cutText(str2(i),3,3,"picture_url","picture_urls")
        str3(15) = cutText(str2(i),3,3,"picture_urls","property_type")
        str3(16) = cutText(str2(i),3,3,"property_type","public_address")
        str3(17) = cutText(str2(i),3,3,"public_address","reviews_count")
        str3(18) = cutText(str2(i),2,2,"reviews_count","star_rating")
        str3(19) = cutText(str2(i),2,2,"star_rating","room_type")
        str3(20) = cutText(str2(i),3,3,"room_type","user")
        str3(21) = cutText(str2(i),3,3,"user","primary_host")
        str3(22) = cutText(str2(i),3,3,"primary_host","coworker_hosted")
        str3(23) = cutText(str2(i),2,2,"coworker_hosted","listing_tags")
        str3(24) = cutText(str2(i),2,3,"listing_tags","pricing_quote")
        str3(25) = cutText(str2(i),3,0,"pricing_quote","viewed_at")
        str3(26) = cutText(str2(i),2,0,"viewed_at","")

        detail20(1) = cutText(str3(21),3,3,"first_name","id")
        detail20(2) = cutText(str3(21),2,2,"id","thumbnail_url")
        detail20(3) = cutText(str3(21),3,3,"thumbnail_url","is_superhost")
        detail20(4) = cutText(str3(21),2,0,"is_superhost","")

        detail21(1) = cutText(str3(22),3,3,"first_name","id")
        detail21(2) = cutText(str3(22),2,2,"id","thumbnail_url")
        detail21(3) = cutText(str3(22),3,3,"thumbnail_url","is_superhost")
        detail21(4) = cutText(str3(22),2,0,"is_superhost","")

        detail25(1) = cutText(str3(25),2,2,"available","can_instant_book")
        detail25(2) = cutText(str3(25),2,2,"can_instant_book","check_in")
        detail25(3) = cutText(str3(25),2,2,"check_in","check_out")
        detail25(4) = cutText(str3(25),2,2,"check_out","guests")
        detail25(5) = cutText(str3(25),2,2,"guests","rate")
        detail25(6) = cutText(str3(25),2,2,"amount","currency")
        detail25(7) = cutText(str3(25),2,3,"currency","rate_type")
        detail25(8) = cutText(str3(25),3,3,"rate_type","is_good_price")
        detail25(9) = cutText(str3(25),2,2,"is_good_price","average_booked_price")
        detail25(10) = cutText(str3(25),2,3,"average_booked_price","")
        For j = 1 to 26
          Debug.print str3(j)
        Next j
        For j = 1 to 4
          Debug.print detail20(j)
        Next j
        For j = 1 to 4
          Debug.print detail21(j)
        Next j
        For j = 1 to 11
          Debug.print detail25(j)
        Next j
        DoEvents
      Next i
    End If
    DoEvents
  Next
End Sub
'============================================='
Function cutText(str As Variant, _
            delStart As Long, _
            delEnd As Long, _
            strStart As String, _
            Optional strEnd As String)
    If IsNull(str) = True Then
    Else
      If Instr(str,strStart) <> 0 Then
        If strEnd = "" Then
          cutText = Mid(str,Instr(str,strStart) + Len(strStart) + delStart,Len(str) - Instr(str,strStart) - Len(strStart) - delStart - delEnd + 1)
        Else
          cutText = Mid(str,Instr(str,strStart) + Len(strStart) + delStart,Instr(str,strEnd) - Instr(str,strStart) - Len(strStart) - delStart - delEnd)
        End If
      End IF
    End If
End Function


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
