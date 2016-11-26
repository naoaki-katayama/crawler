Sub Gmail1()

Dim objIE  As InternetExplorer
Dim URLstr As String
URLstr = "https://mail.google.com/mail/u/0/#inbox"
Call ieView(objIE, URLstr)

Dim str As String
n = 0
For i = 0 to objIE.document.all.tags("td").Length - 1
  'Debug.Print objIE.document.all.tags("td").Length
  'Debug.Print objIE.document.tag("td")(i).outerHTML
   str =  objIE.document.getElementsByTagName("td")(0).outerHTML
   If InStr(str, "yX xY ") > 0 Then
    Debug.Print str
    n = n + 1
   End If
'  For Each objTag In objIE.document.getElementsByTagName(tagName)
'    If InStr(objTag.outerHTML, "td") > 0 Then
'      Debug.Print objTag.outerHTML
'    End If
'  Next

Next i
Debug.Print n
For Each objTag In objIE.document.getElementsByTagName("td")
  Call tagClick(objIE,"td","yX xY ")
  Stop
  objIE.GoBack
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
