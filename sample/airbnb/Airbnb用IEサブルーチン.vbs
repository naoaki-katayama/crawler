#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

''１．ieView＝指定したURLをIE表示する'http://www.vba-ie.net/code/ieview2.html
'ieCheckが必須
''=================================================================

Sub ieView(objIE As InternetExplorer, _
           urlName As String, _
           Optional viewFlg As Boolean = True, _
           Optional ieTop As Integer = 0, _
           Optional ieLeft As Integer = 0, _
           Optional ieWidth As Integer = 600, _
           Optional ieHeight As Integer = 800)

  'IE(InternetExplorer)のオブジェクトを作成する
    Set objIE = CreateObject("InternetExplorer.Application")

    With objIE

    'IE(InternetExplorer)を表示・非表示
        .Visible = viewFlg

        .Top = ieTop 'Y位置
        .Left = ieLeft 'X位置
        .Width = ieWidth '幅
        .Height = ieHeight '高さ

    '指定したURLのページを表示する
        .navigate urlName

    End With

  'IE(InternetExplorer)が完全表示されるまで待機
    Call ieCheck(objIE)

End Sub

''２．ieCheck：IEが完全に読み込まれるまで処理を待機'http://www.vba-ie.net/code/iecheck.html
''=================================================================
Sub ieCheck(objIE As InternetExplorer)

 Dim timeOut As Date

 '完全にページが表示されるまで待機する
 timeOut = Now + TimeSerial(0, 0, 20)

 Do While objIE.Busy = True Or objIE.ReadyState <> 4
  DoEvents
  Sleep 1
  If Now > timeOut Then
   objIE.Refresh
   timeOut = Now + TimeSerial(0, 0, 20)
   End If
 Loop

 timeOut = Now + TimeSerial(0, 0, 20)

 Do While objIE.document.ReadyState <> "complete"
  DoEvents
  Sleep 1
  If Now > timeOut Then
   objIE.Refresh
   timeOut = Now + TimeSerial(0, 0, 20)
  End If
  Loop

End Sub

''３．ieNavi:IEを開いた状態で別のURLを表示'http://www.vba-ie.net/ie/navigate.html'
''=================================================================
Sub ieNavi(objIE As InternetExplorer, _
      urlName As String)

 '指定したURLをIE(InternetExplorer)で表示
 objIE.Navigate urlName

 'IE(InternetExplorer)が完全表示されるまで待機
 Call ieCheck(objIE)

End Sub
