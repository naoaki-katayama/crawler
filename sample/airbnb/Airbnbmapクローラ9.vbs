Sub Airbnbmapクローラ9 ()

  Dim objIE  As InternetExplorer
  Dim URLstr As String
  Dim htmlall As String
  Dim SearchStr As String
  Dim pageCount As Long
  Dim ws(1 to 2) As Worksheet
  Set ws(1) = Worksheets("data")
  Set ws(2) = Worksheets("変数定義")
  Dim ieCount As Integer
  ieCount = 0
  Dim i As Long, j As Long, k As Long, l As Long, Digit As Long


  Dim CheckX(1 to 13) As Long
  Dim CheckY(1 to 13) As Long
  Dim strX1 As String
  Dim strX2 As String
  Dim strY1 As String
  Dim strY2 As String
  Dim Ad As String

  For i = 1 to 13
   CheckX(i) = 0
   CheckY(i) = 0
  Next i

  Digit = 0


  For i = 1 to ws(2).Cells(1,1).End(xlDown).Row
    ws(1).Cells(1,1) = i
    If ws(2).Cells(i,5) = 1 Then
      Do
        strX1 = ""
        strY1 = ""
        strX2 = ""
        strY2 = ""
        For j = 1 to 13
         strX1 = strX1 & CheckX(j)
         strY1 = strY1 & CheckY(j)
         If j = Digit + 1 Then
           If CheckX(j) = 9 Then
             k = j
             Do
              k = k - 1
             Loop While CheckX(k) = 9
             K = k + 1
             Ad = ""
             For l = 1 to k
              Ad = Ad & "0"
             Next l
             strX2 = Mid(strX2,1,Len(strX2) - k) & CheckX(j - k) + 1 & Ad
           Else
             strX2 = strX2 & CheckX(j) + 1
           End If

           If CheckY(j) = 9 Then
             k = j
             Do
               k = k - 1
             Loop While CheckY(k) = 9
             K = k + 1
             Ad = ""
             For l = 1 to k
               Ad = Ad & "0"
             Next l
             strY2 = Mid(strY2,1,Len(strY2) - k) & CheckY(j - k) + 1 & Ad
           Else
             strY2 = strY2 & CheckY(j) + 1
           End If
         Else
           strX2 = strX2 & CheckX(j)
           strY2 = strY2 & CheckY(j)
         End If
        Next j

        URLstr = "https://www.airbnb.jp/s/%E5%A4%A7%E9%98%AA%E5%B8%82?page=1&source=map&airbnb_plus_only=false" _
          & "&sw_lat=" & ws(2).Cells(i,1) & "." & strX1 _
          & "&sw_lng=" & ws(2).Cells(i,2) & "." & strY1 _
          & "&ne_lat=" & ws(2).Cells(i,1) & "." & strX2 _
          & "&ne_lng=" & ws(2).Cells(i,2) & "." & strY2 _
          & "&search_by_map=True"
        Debug.Print "南：" & ws(2).Cells(i, 1) & "." & strX1
        Debug.Print "西：" & ws(2).Cells(i, 2) & "." & strY1
        Debug.Print "北：" & ws(2).Cells(i, 1) & "." & strX2
        Debug.Print "東：" & ws(2).Cells(i, 2) & "." & strY2
        For  j = 1 to Digit
          Debug.Print "checkX(" & j & "): " & checkX(j)
          Debug.Print "checkY(" & j & "): " & checkY(j)
        Next j
        Stop
        If ieCount = 0 Then
          Call ieView(objIE, URLstr)
          ieCount = 1'Call IE
        Else
          Call ieNavi(objIE, URLstr)
        End If

        Application.Wait (DateAdd("s", 2, Now))

        If InStr(objIE.document.all(0).outerHTML,"検索結果300+件") <> 0 Then'300件+のケース
          Debug.Print "over300"
          Digit = Digit + 1
        Else
          If InStr(objIE.document.all(0).outerHTML,">全0件</span>") <> 0 Then'0件のケース
            Debug.Print "under0"
          Else'1~300件のケース
            pageCount = 2
            Do
              'スクレイピング実施
              Call   Airbnbmaptextscraping(objIE,ws(1))
              SearchStr = "page=" & pageCount
              Debug.Print "pageCount:" & pageCount
              'Stop
              Call tagClick2(objIE,"a",SearchStr,pageCount)
              Application.Wait (DateAdd("s", 3, Now))

              pageCount = pageCount + 1
            Loop while pageCount > 1
            ws(1).Cells.WrapText = False
            '処理実行
          End If

          If Digit = 0 Then
            Exit Do
          End if

          CheckX(Digit) = CheckX(Digit) + 1

          Do
            If CheckX(Digit) = 10 Then
              CheckX(Digit) = 0
              CheckY(Digit) = CheckY(Digit) + 1
            End If
            If CheckY(Digit) = 10 Then
              CheckY(Digit) = 0
              Digit = Digit - 1
              If Digit = 0 Then
                Exit Do
              End If
              CheckX(Digit) = CheckX(Digit) + 1
            End If
          Loop While CheckX(Digit) = 10 or CheckY(Digit) = 10
        End If

        ActiveWorkbook.Save

      Loop While Digit > 0
    End If
  Next i
  objIE.Quit

End Sub


'==================================================='
Sub Airbnbmaptextscraping(objIE As InternetExplorer, _
                          ws As Worksheet)
  Dim str As String
  Dim str2 As Variant
  Dim str3(1 to 26) As String
  Dim detail20(1 to 4) As String
  Dim detail21(1 to 4) As String
  Dim detail25(1 to 11) As String
  Dim i As Long, j As Long
  Dim inputRow As Long

  For Each objTag In objIE.document.getElementsByTagName("script")
    str = objTag.outerHTML
    If InStr(objTag.outerHTML, "data-hypernova-key") > 0 Then
      str = Mid(str,Instr(str,"""listing"""),Instr(str,"""metadata""") - Instr(str,"""listing"""))
      str = Mid(str,1,Instr(str,"}]"))
      str2 = split(str,"},{")
      'Debug.Print "UBound(str2):" & UBound(str2)
      For i = LBound(str2) To UBound(str2)
        'Stop
        str3(1) = cutText(str2(i),1,1,"""bedrooms""","""beds""")
        str3(2) = cutText(str2(i),1,1,"""beds""","""airbnb_plus_enabled""")
        str3(3) = cutText(str2(i),1,1,"""airbnb_plus_enabled""","""extra_host_languages""")
        str3(4) = cutText(str2(i),2,3,"""extra_host_languages""",ws.Cells(1,6).Value)
        str3(5) = cutText(str2(i),0,1,ws.Cells(1,6).Value,"""instant_bookable""")
        str3(6) = cutText(str2(i),1,1,"""instant_bookable""","""is_business_travel_ready""")
        str3(7) = cutText(str2(i),1,1,"""is_business_travel_ready""","""is_new_listing""")
        str3(8) = cutText(str2(i),1,1,"""is_new_listing""","""lat""")
        str3(9) = cutText(str2(i),1,1,"""lat""","""lng""")
        str3(10) = cutText(str2(i),1,1,"""lng""","""name""")
        str3(11) = cutText(str2(i),2,2,"""name""","""person_capacity""")
        str3(12) = cutText(str2(i),1,1,"""person_capacity""","""picture_count""")
        str3(13) = cutText(str2(i),1,1,"""picture_count""","""picture_url""")
        str3(14) = cutText(str2(i),2,2,"""picture_url""","""picture_urls""")
        str3(15) = cutText(str2(i),2,2,"""picture_urls""","""property_type""")
        str3(16) = cutText(str2(i),2,2,"""property_type""","""public_address""")
        str3(17) = cutText(str2(i),2,2,"""public_address""","""reviews_count""")
        str3(18) = cutText(str2(i),1,1,"""reviews_count""","""star_rating""")
        str3(19) = cutText(str2(i),1,1,"""star_rating""","""room_type""")
        str3(20) = cutText(str2(i),2,2,"""room_type""","""user""")
        str3(21) = cutText(str2(i),2,2,"""user""","""primary_host""")
        str3(22) = cutText(str2(i),2,2,"""primary_host""","""coworker_hosted""")
        str3(23) = cutText(str2(i),1,1,"""coworker_hosted""","""listing_tags""")
        str3(24) = cutText(str2(i),1,2,"""listing_tags""","""pricing_quote""")
        str3(25) = cutText(str2(i),2,-1,"""pricing_quote""","""viewed_at""")
        str3(26) = cutText(str2(i),1,0,"""viewed_at""","")

        detail20(1) = cutText(str3(21),2,2,"""first_name""","""id""")
        detail20(2) = cutText(str3(21),1,1,"""id""","""thumbnail_url""")
        detail20(3) = cutText(str3(21),2,2,"""thumbnail_url""","""is_superhost""")
        detail20(4) = cutText(str3(21),1,0,"""is_superhost""","")

        detail21(1) = cutText(str3(22),2,2,"""first_name""","""id""")
        detail21(2) = cutText(str3(22),1,1,"""id""","""thumbnail_url""")
        detail21(3) = cutText(str3(22),2,2,"""thumbnail_url""","""is_superhost""")
        detail21(4) = cutText(str3(22),1,0,"""is_superhost""","")

        detail25(1) = cutText(str3(25),1,1,"""available""","""can_instant_book""")
        detail25(2) = cutText(str3(25),1,1,"""can_instant_book""","""check_in""")
        detail25(3) = cutText(str3(25),1,1,"""check_in""","""check_out""")
        detail25(4) = cutText(str3(25),1,1,"""check_out""","""guests""")
        detail25(5) = cutText(str3(25),1,1,"""guests""","""rate""")
        detail25(6) = cutText(str3(25),1,1,"""amount""","""currency""")
        detail25(7) = cutText(str3(25),1,2,"""currency""","""rate_type""")
        detail25(8) = cutText(str3(25),2,2,"""rate_type""","""is_good_price""")
        detail25(9) = cutText(str3(25),1,1,"""is_good_price""","""average_booked_price""")
        detail25(10) = cutText(str3(25),1,3,"""average_booked_price""","")

        inputRow = ws.Cells(1,1).End(xlDown).Row + 1
        ws.Cells(inputRow,1) = inputRow - 2


        ws.Cells(inputRow,2) = str3(1) 'bedrooms
        ws.Cells(inputRow,3) = str3(2) 'beds
        ws.Cells(inputRow,4) = str3(3) 'airbnb_plus_enabled
        ws.Cells(inputRow,5) = str3(4) 'extra_host_languages
        ws.Cells(inputRow,6) = str3(5) 'id
        ws.Cells(inputRow,7) = str3(6) 'instant_bookable
        ws.Cells(inputRow,8) = str3(7) 'is_business_travel_ready
        ws.Cells(inputRow,9) = str3(8) 'is_new_listing
        ws.Cells(inputRow,10) = str3(9) 'lat
        ws.Cells(inputRow,11) = str3(10) 'lng
        ws.Cells(inputRow,12) = str3(11) 'name
        ws.Cells(inputRow,13) = str3(12) 'person_capacity
        ws.Cells(inputRow,14) = str3(13) 'picture_count
        ws.Cells(inputRow,15) = str3(14) 'picture_url
        ws.Cells(inputRow,16) = str3(15) 'picture_urls
        ws.Cells(inputRow,17) = str3(16) 'property_type
        ws.Cells(inputRow,18) = str3(17) 'public_address
        ws.Cells(inputRow,19) = str3(18) 'reviews_count
        ws.Cells(inputRow,20) = str3(19) 'star_rating
        ws.Cells(inputRow,21) = str3(20) 'room_type
        ws.Cells(inputRow,22) = str3(23) 'coworker_hosted
        ws.Cells(inputRow,23) = str3(24) 'listing_tags
        ws.Cells(inputRow,24) = str3(26) 'viewed_at
        ws.Cells(inputRow,25) = detail20(1) 'user>first_name
        ws.Cells(inputRow,26) = detail20(2) 'user>id
        ws.Cells(inputRow,27) = detail20(3) 'user>thumbnail_url
        ws.Cells(inputRow,28) = detail20(4) 'user>is_superhost
        ws.Cells(inputRow,29) = detail21(1) 'primary_host>first_name
        ws.Cells(inputRow,30) = detail21(2) 'primary_host>id
        ws.Cells(inputRow,31) = detail21(3) 'primary_host>thumbnail_url
        ws.Cells(inputRow,32) = detail21(4) 'primary_host>is_superhost
        ws.Cells(inputRow,33) = detail25(1) 'pricing_quote>available
        ws.Cells(inputRow,34) = detail25(2) 'pricing_quote>can_instant_book
        ws.Cells(inputRow,35) = detail25(3) 'pricing_quote>check_in
        ws.Cells(inputRow,36) = detail25(4) 'pricing_quote>check_out
        ws.Cells(inputRow,37) = detail25(5) 'pricing_quote>guests
        ws.Cells(inputRow,38) = detail25(6) 'pricing_quote>amount
        ws.Cells(inputRow,39) = detail25(7) 'pricing_quote>currency
        ws.Cells(inputRow,40) = detail25(8) 'pricing_quote>rate_type
        ws.Cells(inputRow,41) = detail25(9) 'pricing_quote>is_good_price
        ws.Cells(inputRow,42) = detail25(10) 'pricing_quote>average_booked_price
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
