Sub calltest()
fileName = "C:\Users\USER\Dropbox\開発\ruby\airbnbtext2.txt"

Call Airbnbwgetしたテキスト変換4(fileName)

End Sub

Sub Airbnbwgetしたテキスト変換4(ByRef fileName As Variant)

'utf-8形式ファイル読み込み
'参考：http://www.hiihah.info/index.php?Excel%EF%BC%9AVBA%EF%BC%9AUTF-8%EF%BC%8FLF%E3%81%AE%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%82%92%E8%AA%AD%E3%81%BF%E8%BE%BC%E3%82%80

Dim rowNo As Integer
Dim readString As String
Dim st As Object

  '  ADODB.Streamの参照URL
  '  http://msdn.microsoft.com/ja-jp/library/cc364272.aspx
  '  http://msdn.microsoft.com/ja-jp/library/cc364273.aspx

Set st = CreateObject("ADODB.Stream") 'ADODB.Stream生成

  '  StreamTypeEnumの仕様
  '　adTypeBinary    1   バイナリ データを表します。
  '  adTypeText  2   既定値です。Charset で指定された文字セットにあるテキスト データを表します。
  '  【参照URL】http://msdn.microsoft.com/ja-jp/library/cc389884.aspx

st.Type = 2  'オブジェクトに保存するデータの種類を文字列型に指定する

'  【参照URL】http://msdn.microsoft.com/ja-jp/library/cc364313.aspx
st.Charset = "utf-8"  '文字コード（Shift_JIS, Unicodeなど）

'  LineSeparatorsEnumの仕様
'  adCR    13  改行復帰を示します。
'  adCRLF  -1  既定値です。改行復帰行送りを示します。
'  adLF    10  行送りを示します。
'  【参照URL】http://msdn.microsoft.com/ja-jp/library/cc389826.aspx
st.LineSeparator = 10 '改行LF（10）

st.Open         'Streamのオープン
st.LoadFromFile (fileName)  'ファイル読み込み


rowNo = 5
'ファイルの終りまでループ
Do While Not st.EOS
  rowNo = rowNo + 1
'  ReadTextの仕様
'　引数、読み込む文字列、もしくはStreamReadEnumの値を指定
'　StreamReadEnumの仕様
'　adReadAll   -1  既定値です。現在の位置から EOS マーカー方向に、すべてのバイトをストリームから読み取ります。これは、バイナリ ストリームに唯一有効な StreamReadEnum 値です (Type は adTypeBinary)。
'　adReadLine  -2  ストリームから次の行を読み取ります (LineSeparator プロパティで指定)。
'　【参照URL】http://msdn.microsoft.com/ja-jp/library/cc364207.aspx

  readString = st.ReadText(-2) 'テキストを1行読み込む。

  Cells(rowNo, 2).Value = readString   '読み込んだ文字列をセルにセットする
Loop

st.Close  'Streamのクローズ
Set st = Nothing

End Sub
