Sub Gmail3()

Dim objIE  As InternetExplorer
Dim URLstr As String
URLstr = "https://mail.google.com/mail/u/0/#inbox"
Dim URL As String
Dim pageCount As Variant
Dim i As Long
Call ieView(objIE, URLstr)
Application.Wait (DateAdd("s", 10, Now))
URL = objIE.document.URL

pageCount = cutText(objIE.document.all(0).outerHTML,1,5,"""優先トレイ""","""in:inbox""")
If pageCount Mod 50 = 0 Then
  pageCount = Int(pageCount / 50)
Else
  pageCount = Int(pageCount / 50) + 1
End If
Debug.Print pageCount

'Stop
For i = 1 To pageCount - 1

  For Each objTag In objIE.document.getElementsByTagName("td")
    If InStr(objTag.outerHTML, "yX xY ") > 0 Then
      Debug.Print URL
      objTag.Click
      Call ieCheck(objIE)
      Debug.Print objIE.document.URL
      'Stop
      If objIE.document.URL = URL Then
      Else
        Call InputText(objIE)
        objIE.GoBack
        Call ieCheck(objIE)
      End If
    End If
  Next
  Application.Wait (DateAdd("s", 2, Now))
  Call tagClick(objIE,"img","amJ T-I-J3")
  Application.Wait (DateAdd("s", 10, Now))
  URL = objIE.document.URL

Next i

End Sub
'==================================================='
Sub inputText(objIE As InternetExplorer)
  Call tagClick3(objIE,"div","kQ hn ")
  Call ieCheck(objIE)
  Call tagClick3(objIE,"div","gE hI")
  Call ieCheck(objIE)
  Call inputText2(objIE,"div","nH hx")
  'Debug.Print objIE.document.all(0).innerText
End Sub
'==================================================='
Sub inputText2(objIE As InternetExplorer, _
             tagName As String, _
             tagStr As String)

  For Each objTag In objIE.document.getElementsByTagName(tagName)
    If InStr(objTag.outerHTML, tagStr) > 0 Then
      Debug.Print objTag.outerHTML
      Exit For
    End If
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
'==================================================='
Sub tagClick3(objIE As InternetExplorer, _
             tagName As String, _
             tagStr As String)

  'タグをクリック
  For Each objTag In objIE.document.getElementsByTagName(tagName)
    If InStr(objTag.outerHTML, tagStr) > 0 Then
      objTag.Click
      Call ieCheck(objIE)
    End If
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
