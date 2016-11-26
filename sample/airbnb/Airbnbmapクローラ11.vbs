Sub Airbnbmapクローラ11 ()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


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
  Dim strX1 As Variant
  Dim strX2 As Variant
  Dim strY1 As Variant
  Dim strY2 As Variant
  'Dim Ad As String

  For i = 1 to 13
   CheckX(i) = 0
   CheckY(i) = 0
  Next i

  Digit = 0


  For i = 1 to ws(2).Cells(1,1).End(xlDown).Row
    ws(1).Cells(1,1) = i
    If ws(2).Cells(i,5) = 1 Then
      Do
        strX1 = 10 ^14
        strY1 = 10 ^14
        For j = 1 to 13
          If CheckX(j) > 0 Then
            strX1 = strX1 + CheckX(j) * 10 ^ (14 - j)
            'Debug.Print strX1
          End If
          If CheckY(j) > 0 Then
            strY1 = strY1 + CheckY(j) * 10 ^ (14 - j)
          End If
        Next j

        strX2 = strX1 + 10 ^ (14 - Digit)
        strY2 = strY1 + 10 ^ (14 - Digit)

        strX1 = Mid(strX1,1,1) & "." & Mid(strX1,2)
        strY1 = Mid(strY1,1,1) & "." & Mid(strY1,2)
        strX2 = Mid(strX2,1,1) & "." & Mid(strX2,2)
        strY2 = Mid(strY2,1,1) & "." & Mid(strY2,2)

        strX1 = ws(2).Cells(i,1) + strX1 - 1
        strY1 = ws(2).Cells(i,2) + strY1 - 1
        strX2 = ws(2).Cells(i,1) + strX2 - 1
        strY2 = ws(2).Cells(i,2) + strY2 - 1


        URLstr = "https://www.airbnb.jp/s/大阪市?page=1&source=map&airbnb_plus_only=false" _
          & "&sw_lat=" & strX1 _
          & "&sw_lng=" & strY1 _
          & "&ne_lat=" & strX2 _
          & "&ne_lng=" & strY2 _
          & "&search_by_map=true"
        Debug.Print URLstr
        Debug.Print "南：" & strX1
        Debug.Print "西：" & strY1
        Debug.Print "北：" & strX2
        Debug.Print "東：" & strY2
        For  j = 1 to Digit
          Debug.Print "checkX(" & j & "): " & checkX(j)
          Debug.Print "checkY(" & j & "): " & checkY(j)
        Next j
        'Stop
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

  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic


End Sub


'==================================================='
Sub Airbnbmaptextscraping(objIE As InternetExplorer, _
                          ws As Worksheet)
  Dim str As String
  Dim str2 As Variant
  Dim str3(1 to 28) As String
  Dim detail21(1 to 4) As String
  Dim detail22(1 to 4) As String
  Dim detail28(1 to 9) As String
  Dim i As Long, j As Long
  Dim inputRow As Long
  Dim ws1 As Worksheet
  Set ws1 = Worksheets("変数定義")

  Dim cut(1 to 49) As String
  For i = 1 to 49
    cut(i) = ws1.Cells(i,11)
  Next i

  For Each objTag In objIE.document.getElementsByTagName("script")
    str = objTag.outerHTML
    If InStr(objTag.outerHTML, "data-hypernova-key") > 0 Then
      str = Mid(str,Instr(str,"""listing""") + 11 ,Instr(str,"""metadata""") - Instr(str,"""listing""") - 12)
      str2 = split(str,ws1.Cells(1,8).value)
      'Debug.Print "UBound(str2):" & UBound(str2)
      For i = LBound(str2) To UBound(str2) - 1
      If i Mod 2 = 1 Then
       GoTo Continue
      End If
        'Stop
        str3(1) = cutText(str2(i),0,1,cut(1),cut(2))
        str3(2) = cutText(str2(i),0,1,cut(2),cut(3))
        str3(3) = cutText(str2(i),0,1,cut(3),cut(4))
        str3(4) = cutText(str2(i),1,2,cut(4),cut(5))
        str3(5) = cutText(str2(i),0,1,cut(5),cut(6))
        str3(6) = cutText(str2(i),0,1,cut(6),cut(7))
        str3(7) = cutText(str2(i),0,1,cut(7),cut(8))
        str3(8) = cutText(str2(i),0,1,cut(8),cut(9))
        str3(9) = cutText(str2(i),0,1,cut(9),cut(10))
        str3(10) = cutText(str2(i),0,1,cut(10),cut(11))
        str3(11) = cutText(str2(i),1,2,cut(11),cut(12))
        str3(12) = cutText(str2(i),0,1,cut(12),cut(13))
        str3(13) = cutText(str2(i),0,1,cut(13),cut(14))
        str3(14) = cutText(str2(i),0,1,cut(14),cut(15))
        str3(15) = cutText(str2(i),1,2,cut(15),cut(16))
        str3(16) = cutText(str2(i),1,2,cut(16),cut(17))
        str3(17) = cutText(str2(i),1,2,cut(17),cut(18))
        str3(18) = cutText(str2(i),0,1,cut(18),cut(19))
        str3(19) = cutText(str2(i),0,1,cut(19),cut(20))
        str3(20) = cutText(str2(i),1,2,cut(20),cut(21))
        str3(21) = cutText(str2(i),0,0,cut(21),cut(22))
        str3(22) = cutText(str2(i),0,0,cut(22),cut(23))
        str3(23) = cutText(str2(i),0,1,cut(23),cut(24))
        str3(24) = cutText(str2(i),1,2,cut(24),cut(25))
        str3(25) = cutText(str2(i),1,2,cut(25),cut(26))
        str3(26) = cutText(str2(i),0,1,cut(26),cut(27))
        str3(27) = cutText(str2(i),2,2,cut(27),cut(28))
        str3(28) = cutText(str2(i),1,0,cut(28),cut(29))

        detail21(1) = cutText(str3(21),1,2,cut(30),cut(31))
        detail21(2) = cutText(str3(21),0,1,cut(31),cut(32))
        detail21(3) = cutText(str3(21),1,2,cut(32),cut(33))
        detail21(4) = cutText(str3(21),0,2,cut(33))

        detail22(1) = cutText(str3(22),1,2,cut(35),cut(36))
        detail22(2) = cutText(str3(22),0,1,cut(36),cut(37))
        detail22(3) = cutText(str3(22),1,2,cut(37),cut(38))
        detail22(4) = cutText(str3(22),0,2,cut(38))

        detail28(1) = cutText(str3(28),0,1,cut(40),cut(41))
        detail28(2) = cutText(str3(28),0,1,cut(41),cut(42))
        detail28(3) = cutText(str3(28),0,1,cut(42),cut(43))
        detail28(4) = cutText(str3(28),0,9,cut(43),cut(44))
        detail28(5) = cutText(str3(28),0,1,cut(44),cut(45))
        detail28(6) = cutText(str3(28),1,3,cut(45),cut(46))
        detail28(7) = cutText(str3(28),1,2,cut(46),cut(47))
        detail28(8) = cutText(str3(28),0,1,cut(47),cut(48))
        detail28(9) = cutText(str3(28),0,2,cut(48))

        inputRow = ws.Cells(1,1).End(xlDown).Row + 1
        ws.Cells(inputRow,1) = inputRow - 2

        For j = 2 to 43
          Select Case j
            Case 2 to 21
              ws.Cells(inputRow,j) = str3(j - 1)
            Case 22 to 26
              ws.Cells(inputRow,j) = str3(j + 1)
            Case 27 to 30
              ws.Cells(inputRow,j) = detail21(j - 26)
            Case 31 to 34
              ws.Cells(inputRow,j) = detail22(j - 30)
            Case 35 to 43
              ws.Cells(inputRow,j) = detail28(j - 34)
          End Select
        Next j
        Continue:
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
