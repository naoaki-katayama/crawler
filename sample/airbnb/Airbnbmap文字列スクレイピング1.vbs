Sub Airbnbmaptextscraping(objIE As InternetExplorer)
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
        detail25(7) = cutText(str3(25),2,2,"currency","rate_type")
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
