Sub テキストファイルをcsv化1()

  Dim buf As String
  Open "C:\Users\尚亮\Dropbox\MKP\6.Airbnb\クローリング\airbnb10112039.txt" For Input As #1
  Do Until EOF(1)
      Line Input #1, buf
      n = n + 1
      Cells(n, 1) = buf
  Loop
  Close #1
End Sub
