Sub Airbnbwgetしたテキスト変換2()

'utf-8形式ファイル読み込み

    Dim buf As String, Target As String
    Target = "C:\Users\USER\Dropbox\開発\ruby\airbnbtext2.txt"
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile Target
        buf = .ReadText
        .Close
        MsgBox buf
    End With
End Sub
