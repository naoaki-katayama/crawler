Sub Airbnbwgetしたテキスト変換3()

'ファイルを読み込むための配列
Dim Arr()
ReDim Preserve Arr(0)

'オブジェクトを作成
Dim txt As Object
Set txt = CreateObject("ADODB.Stream")

'オブジェクトに保存するデータの種類を文字列型に指定する
txt.Type = adTypeText
'文字列型のオブジェクトの文字コードを指定する
txt.Charset = "UTF-8"

'オブジェクトのインスタンスを作成
txt.Open

'ファイルからデータを読み込む
txt.LoadFromFile ("C:\Users\USER\Dropbox\開発\ruby\airbnbtext2.txt")

'最終行までループする
Do While Not txt.EOS
    '次の行を読み取る
    Arr(UBound(Arr)) = txt.ReadText(adReadLine)
    ReDim Preserve Arr(UBound(Arr) + 1)
    Debug.Print Arr(UBound(Arr))
Loop

'オブジェクトを閉じる
txt.Close

'メモリからオブジェクトを削除する
Set txt = Nothing

End Sub
