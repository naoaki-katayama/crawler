Sub textfileSave()

Dim objIE  As InternetExplorer
Dim URLstr As String
Dim str As String
Dim datFile As String
datFile = ActiveWorkbook.Path & "\data.txt"

URLstr = "https://www.airbnb.jp/rooms/1303278"
Call ieView(objIE, URLstr)
Application.Wait (DateAdd("s", 2, Now))
str = objIE.document.all(0).outerHTML

Open datFile For Output As #1
Print #1, str
Close #1

End Sub
